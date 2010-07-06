VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPointer 
   Caption         =   "Zuri's Pointer"
   ClientHeight    =   7005
   ClientLeft      =   165
   ClientTop       =   1155
   ClientWidth     =   9660
   Icon            =   "Pointer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   9660
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDescription 
      BackColor       =   &H00E0F0E0&
      Height          =   855
      Left            =   840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Top             =   5760
      Width           =   2895
   End
   Begin MSFlexGridLib.MSFlexGrid grdNonSpec 
      Height          =   1455
      Left            =   1560
      TabIndex        =   15
      Top             =   4080
      Visible         =   0   'False
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   2566
      _Version        =   393216
      Rows            =   6
      Cols            =   0
      FixedCols       =   0
      BackColor       =   14741728
      BackColorBkg    =   -2147483633
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.TextBox txtAttributeAdditional 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6480
      TabIndex        =   11
      Top             =   5760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSpec 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   7440
      TabIndex        =   10
      Top             =   5760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAttributeBuff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5520
      TabIndex        =   9
      Top             =   5760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame fraSkills 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   9375
      Begin VB.CheckBox chkSpecPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Allocate by percent"
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   135
         Width           =   1815
      End
      Begin VB.TextBox txtPointsAcquired 
         BackColor       =   &H00E0F0E0&
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "0"
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox txtPointsSpent 
         BackColor       =   &H00E0F0E0&
         Height          =   285
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "0"
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox txtPointsAvailable 
         BackColor       =   &H00E0F0E0&
         Height          =   285
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "0"
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lbl 
         Caption         =   "Points Acquired"
         Height          =   255
         Index           =   3
         Left            =   3600
         TabIndex        =   8
         Top             =   135
         Width           =   1335
      End
      Begin VB.Label lbl 
         Caption         =   "Points Spent"
         Height          =   255
         Index           =   2
         Left            =   5640
         TabIndex        =   6
         Top             =   135
         Width           =   1335
      End
      Begin VB.Label lbl 
         Caption         =   "Points Available"
         Height          =   255
         Index           =   0
         Left            =   7440
         TabIndex        =   4
         Top             =   135
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdAttributes 
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   2566
      _Version        =   393216
      Rows            =   6
      Cols            =   9
      BackColor       =   14741728
      BackColorBkg    =   -2147483633
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Frame fraCharacter 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      Begin VB.ComboBox cmbName 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   105
         Width           =   2415
      End
      Begin VB.ComboBox cmbRealm 
         Height          =   315
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   105
         Width           =   1335
      End
      Begin VB.ComboBox cmbRace 
         Height          =   315
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   105
         Width           =   1335
      End
      Begin VB.ComboBox cmbClass 
         Height          =   315
         Left            =   6600
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   105
         Width           =   1335
      End
      Begin VB.TextBox txtLevel 
         BackColor       =   &H00E0F0E0&
         Height          =   285
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "5"
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lbl 
         Caption         =   "Level"
         Height          =   255
         Index           =   1
         Left            =   8160
         TabIndex        =   20
         Top             =   135
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdSpec 
      Height          =   1455
      Left            =   480
      TabIndex        =   14
      Top             =   3600
      Visible         =   0   'False
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   2566
      _Version        =   393216
      Rows            =   6
      Cols            =   9
      BackColor       =   14741728
      BackColorBkg    =   -2147483633
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin MSComctlLib.TabStrip tabGrids 
      Height          =   1935
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   3413
      MultiRow        =   -1  'True
      Style           =   1
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Attributes"
            Key             =   "Attributes"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Trainable Lines"
            Key             =   "Trainable Lines"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Level Based Lines"
            Key             =   "Level Based Lines"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuCharacter 
      Caption         =   "&Character"
      Begin VB.Menu mnuCharacterNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuCharacterRename 
         Caption         =   "&Rename"
      End
      Begin VB.Menu mnuCharacterCopy 
         Caption         =   "Make &copy"
      End
      Begin VB.Menu mnuCharacterDelete 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu mnuInternet 
      Caption         =   "&Internet"
      Begin VB.Menu mnuInternetUpdate 
         Caption         =   "Get &Update"
      End
      Begin VB.Menu mnuInternetPointer 
         Caption         =   "&Pointer"
      End
      Begin VB.Menu mnuIntenetShadows 
         Caption         =   "&Shadows"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmPointer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintCellIndex As Integer
Private mintCurGrid As Integer

Private Sub chkSpecPercent_Click()
    Toon.SpecByPercent = chkSpecPercent = vbChecked
    SetSpecGrid
End Sub

Private Sub cmbClass_Click()
' Load specific class data

    SystemWait
    
    If (cmbClass.List(cmbClass.ListIndex) <> vbNullString) And (cmbClass.List(cmbClass.ListIndex) <> "Class") Then
        Me.Refresh
        frmProgress.Progress "Loading " & cmbClass.List(cmbClass.ListIndex) & " Data"
        frmProgress.Move Me.Left + (Me.Width - frmProgress.Width) / 2, Me.Top + (Me.Height - frmProgress.Height) / 2
        frmProgress.Show
        frmProgress.Refresh
        ' Set the Toon class
        Toon.Class = cmbClass.List(cmbClass.ListIndex)
        RemoveListItem cmbClass, "Class"
        ' Set the attributes in the grid
        SetAttributeGrid
        SetPointsAcquired
        SetPointsSpent
        SetLevel
        SetSpecGrid
        SetNonSpecGrid
        UnlockLeveltxt
    End If
    Unload frmProgress
    Me.SetFocus
    SystemContinue
    
End Sub


Private Sub cmbName_Click()
    If (cmbRace.ListIndex <> -1) And (cmbName.List(cmbName.ListIndex) <> Toon.CharName) Then
        Toon.SaveToon
        Toon.NewToon
        Toon.CharName = cmbName.List(cmbName.ListIndex)
        Toon.LoadToon
    End If
End Sub

Private Sub cmbRace_Click()
'Load the relevant classes and race data
    
    If (cmbRace.ListIndex <> -1) And (cmbRace.List(cmbRace.ListIndex) <> "Race") Then
        Toon.Race = cmbRace.List(cmbRace.ListIndex)
        RemoveListItem cmbRace, "Race"
        LoadClasses cmbRace.List(cmbRace.ListIndex)
        SetLevel
        SetAttributeGrid
        SetSpecGrid
        SetNonSpecGrid
        LockLeveltxt
    End If
    
End Sub


Private Sub cmbRealm_Click()
'Load the relevant classes
    
    If (cmbRealm.ListIndex <> -1) And (cmbRealm.List(cmbRealm.ListIndex) <> "Realm") Then
        Toon.Realm = cmbRealm.List(cmbRealm.ListIndex)
        RemoveListItem cmbRealm, "Realm"
        LoadRaces Toon.Realm
        SetLevel
        SetAttributeGrid
        SetSpecGrid
        SetNonSpecGrid
        LockLeveltxt
    End If
    
End Sub

Private Sub Form_Load()

    mintCurGrid = 1

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveNames
    Toon.SaveToon
    CleanUp
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    ' Size tab control
    tabGrids.Move 10 * Screen.TwipsPerPixelX, TABSTOP * Screen.TwipsPerPixelY, Me.ScaleWidth - (20 * Screen.TwipsPerPixelX), Me.ScaleHeight - ((TABSTOP + DESCRIPTIONTOP - 10) * Screen.TwipsPerPixelY)
    ' Size grids to the tab control
    grdAttributes.Move tabGrids.ClientLeft, tabGrids.ClientTop, tabGrids.ClientWidth, tabGrids.ClientHeight
    grdSpec.Move tabGrids.ClientLeft, tabGrids.ClientTop, tabGrids.ClientWidth, tabGrids.ClientHeight
    grdNonSpec.Move tabGrids.ClientLeft, tabGrids.ClientTop, tabGrids.ClientWidth, tabGrids.ClientHeight
    ' Size the description label
    txtDescription.Move 10 * Screen.TwipsPerPixelX, Me.ScaleHeight - ((DESCRIPTIONTOP - 10) * Screen.TwipsPerPixelY), Me.ScaleWidth - (20 * Screen.TwipsPerPixelX), ((DESCRIPTIONTOP - 10) * Screen.TwipsPerPixelY)
    On Error GoTo 0
    
End Sub

Private Sub grdAttributes_Click()
Dim strDescription As String

    ' Bail if race not set
    If Toon.Race = vbNullString Then Exit Sub
    
    ' If an additional or buff cell is selected
    With grdAttributes
        ' Ensure that the current cell is the same as the 'mouse cell'
        .Row = .MouseRow
        .Col = .MouseCol
        If .Col > 0 Then
            strDescription = ProfileGetItem("Attributes", "Attribute " & .Col & " Description", "No description", App.Path & DATAFILE)
            Select Case .MouseRow
                Case ATTRIBUTEADDITIONAL
                    mintCellIndex = ATTRIBUTEADDITIONAL * .Cols + .MouseCol
                    txtAttributeAdditional.Move .Left + .CellLeft, .Top + .CellTop + (1 * Screen.TwipsPerPixelY), .CellWidth - (1 * Screen.TwipsPerPixelX), .CellHeight - (1 * Screen.TwipsPerPixelY)
                    txtAttributeAdditional.Text = .TextArray(mintCellIndex)
                    txtAttributeAdditional.Visible = True
                    txtAttributeAdditional.SetFocus
                Case ATTRIBUTEBUFF
                    mintCellIndex = ATTRIBUTEBUFF * .Cols + .MouseCol
                    txtAttributeBuff.Move .Left + .CellLeft, .Top + .CellTop + (1 * Screen.TwipsPerPixelY), .CellWidth - (1 * Screen.TwipsPerPixelX), .CellHeight - (1 * Screen.TwipsPerPixelY)
                    txtAttributeBuff.Text = .TextArray(mintCellIndex)
                    txtAttributeBuff.Visible = True
                    txtAttributeBuff.SetFocus
                Case Else
            End Select
        Else
            strDescription = vbNullString
        End If
        txtDescription.Text = strDescription

    End With

End Sub

Private Sub grdNonSpec_Click()
Dim strDescription As String

    ' Bail if class not set
    If Toon.Class = vbNullString Then Exit Sub
    
    ' If a level cell is selected
    With grdNonSpec
        ' Ensure that the current cell is the same as the 'mouse cell'
        .Row = .MouseRow
        .Col = .MouseCol
        If (.Col > 0) And (.Col < .Cols - 2) Then
            strDescription = Toon.NonSpecLine(.Col).Description
        Else
            strDescription = vbNullString
        End If
        Select Case .Col
            Case .Cols - 2 ' Armor
                Select Case .Row
                    Case 0 ' Caption
                    Case Else
                        ' Try and load a description
                        On Error Resume Next
                        'If there is no spec line in this column, bail
                        If .TextArray(.Col) = vbNullString Then Exit Sub
                        ' If there is no text in the cell then bail
                        If .TextArray((.Row * .Cols) + .Col) = vbNullString Then Exit Sub
                        strDescription = Toon.ArmorLine(.Row).Description
                        On Error GoTo 0
                End Select
            Case .Cols - 1 ' Other
                Select Case .Row
                    Case 0 ' Caption
                    Case Else
                        ' Try and load a description
                        On Error Resume Next
                        'If there is no spec line in this column, bail
                        If .TextArray(.Col) = vbNullString Then Exit Sub
                        ' If there is no text in the cell then bail
                        If .TextArray((.Row * .Cols) + .Col) = vbNullString Then Exit Sub
                        strDescription = Toon.OtherLine(.Row).Description
                        On Error GoTo 0
                End Select
            Case Else ' Non spec line
                Select Case .Row
                    Case 0 ' Caption
                    Case Else
                        ' Try and load a description
                        On Error Resume Next
                        'If there is no spec line in this column, bail
                        If .TextArray(.Col) = vbNullString Then Exit Sub
                        ' If there is no text in the cell then bail
                        If .TextArray((.Row * .Cols) + .Col) = vbNullString Then Exit Sub
                        strDescription = Toon.NonSpecLine(.Col + 1).Item(.Row).Description
                        On Error GoTo 0
                End Select
            End Select
    End With
    
    txtDescription.Text = strDescription
End Sub

Private Sub grdSpec_Click()
Dim strDescription As String

    ' Bail if class not set
    If Toon.Class = vbNullString Then Exit Sub
    
    ' If a level cell is selected
    With grdSpec
        ' Ensure that the current cell is the same as the 'mouse cell'
        .Row = .MouseRow
        .Col = .MouseCol
        If .Col > 0 Then
            strDescription = Toon.SpecLine(.Col).Description
        Else
            strDescription = vbNullString
        End If
        Select Case .Row
            Case SPECCAPTION
            Case SPECPERCENT
                ' If We are not sllocating by percent then bail
                If Not Toon.SpecByPercent Then Exit Sub
                'If there is no spec line in this column, bail
                If .TextArray(.Col) = vbNullString Then Exit Sub
                strDescription = "Percentage to spec " & Toon.SpecLine(.Col).LineName & vbCrLf & strDescription
               mintCellIndex = SPECPERCENT * .Cols + .Col
                txtSpec.Move .Left + .CellLeft, .Top + .CellTop + (1 * Screen.TwipsPerPixelY), .CellWidth - (1 * Screen.TwipsPerPixelX), .CellHeight - (1 * Screen.TwipsPerPixelY)
                txtSpec.Text = .TextArray(mintCellIndex)
                txtSpec.Visible = True
                txtSpec.SetFocus
            Case SPECPOINTS
                ' If We are allocating by percent then bail
                If Toon.SpecByPercent Then Exit Sub
                'If there is no spec line in this column, bail
                If .TextArray(.Col) = vbNullString Then Exit Sub
                strDescription = "Points allocated to " & Toon.SpecLine(.Col).LineName & vbCrLf & strDescription
               mintCellIndex = SPECPOINTS * .Cols + .Col
                txtSpec.Move .Left + .CellLeft, .Top + .CellTop + (1 * Screen.TwipsPerPixelY), .CellWidth - (1 * Screen.TwipsPerPixelX), .CellHeight - (1 * Screen.TwipsPerPixelY)
                txtSpec.Text = .TextArray(mintCellIndex)
                txtSpec.Visible = True
                txtSpec.SetFocus
            Case SPECLEVEL
                ' If We are allocating by percent then bail
                If Toon.SpecByPercent Then Exit Sub
                'If there is no spec line in this column, bail
                If .TextArray(.Col) = vbNullString Then Exit Sub
                strDescription = "Level in " & Toon.SpecLine(.Col).LineName & vbCrLf & strDescription
                mintCellIndex = SPECLEVEL * .Cols + .Col
                txtSpec.Move .Left + .CellLeft, .Top + .CellTop + (1 * Screen.TwipsPerPixelY), .CellWidth - (1 * Screen.TwipsPerPixelX), .CellHeight - (1 * Screen.TwipsPerPixelY)
                txtSpec.Text = .TextArray(mintCellIndex)
                txtSpec.Visible = True
                txtSpec.SetFocus
            Case Else
                ' Try and load a description
                On Error Resume Next
                'If there is no spec line in this column, bail
                If .TextArray(.Col) = vbNullString Then Exit Sub
                ' If there is no text in the cell then bail
                If .TextArray((.Row * .Cols) + .Col) = vbNullString Then Exit Sub
                strDescription = Toon.SpecLine(.Col).Item(.Row - SPECSTYLES + 1).Description
                On Error GoTo 0
        End Select
    End With
    
    txtDescription.Text = strDescription

End Sub

Private Sub grdSpec_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Try and set level if right mouse button is pressed
Dim intLevelRequired As Integer
Dim intLine As Integer

    ' Bail if class not set
    If (Toon.Class = vbNullString) Or (Toon.Class = "Class") Then Exit Sub
    
    ' Bail if not the right mouse button
    If Button <> vbRightButton Then Exit Sub
    
    ' If a level cell is selected
    With grdSpec
        ' Ensure that the current cell is the same as the 'mouse cell'
        .Row = .MouseRow
        .Col = .MouseCol
        If .Row > SPECLEVEL Then
            ' Try and get the level
            On Error Resume Next
            'If there is no spec line in this column, bail
            If .TextArray(.Col) = vbNullString Then Exit Sub
            ' If there is no text in the cell then bail
            If .TextArray((.Row * .Cols) + .Col) = vbNullString Then Exit Sub
            intLevelRequired = Toon.SpecLine(.Col).Item(.Row - SPECSTYLES + 1).Level
            On Error GoTo 0
            
            ' Try and set the level
            With grdSpec
                 mintCellIndex = SPECPERCENT * .Cols + .Col
                 intLine = mintCellIndex Mod .Cols
                 Toon.LineLevel(intLine) = intLevelRequired
            
                .TextArray((SPECPERCENT * .Cols) + intLine) = Toon.LinePercent(intLine)
                .TextArray((SPECPOINTS * .Cols) + intLine) = Toon.LinePoints(intLine)
                .TextArray((SPECLEVEL * .Cols) + intLine) = Toon.LineLevel(intLine)
            End With
        
            SetSpecPoints
            SetAvailableSpecStyles intLine
            
        End If
    End With

End Sub

Private Sub grdSpec_Scroll()
' Ensure that the the spec text box looses focus
If TypeOf Me.ActiveControl Is TextBox Then grdSpec.SetFocus

End Sub

Private Sub mnuCharacterCopy_Click()
' Copy the current char
Dim strNewName As String
Dim i As Integer
Dim bDuplicate As Boolean

    ' Get name
    strNewName = InputBox("Please enter a new name to copy " & cmbName.Text & " to.")
    ' Check for duplicates
    For i = 0 To cmbName.ListCount - 1
        If UCase(cmbName.List(i)) = UCase(strNewName) Then
            If MsgBox(strNewName & " already in use, overwrite?", vbYesNo) <> vbYes Then
                Exit Sub
            Else
                bDuplicate = True
            End If
            Exit For
        End If
    Next i
    ' Add to combo
    If Not bDuplicate Then cmbName.AddItem strNewName
    ' Set Name
    Toon.CharName = strNewName
    ' Load interface
    LoadFormFromChar

End Sub

Private Sub mnuCharacterDelete_Click()
' Delete the current char
Dim strName As String
Dim i As Integer

    strName = cmbName.Text
    If MsgBox("Delete " & strName & "?", vbYesNo) = vbYes Then
        ' Delete the string from the combo
        For i = cmbName.ListCount - 1 To 0 Step -1
            If cmbName.List(i) = strName Then cmbName.RemoveItem i
        Next i
        ' If the combo is empty
        If cmbName.ListCount = 0 Then
            ' Add a new char called 'No Name'
            Toon.NewToon
            ' Set Name
            Toon.CharName = "No Name"
        Else
            ' Otherwise Select the first char and load it
            Toon.CharName = cmbName.List(0)
            Toon.LoadToon
        End If
        ' Load interface
        LoadFormFromChar

        ' Delete the char from the data file
        DeleteCharData strName
    End If
    
End Sub

Private Sub mnuCharacterNew_Click()
' Create a new character
Dim strNewName As String
Dim i As Integer
Dim bDuplicate As Boolean

    ' Get name
    strNewName = InputBox("Please enter a name")
    If strNewName = vbNullString Then Exit Sub
    If strNewName = cmbName.List(cmbName.ListIndex) Then Exit Sub
    ' Check for duplicates
    For i = 0 To cmbName.ListCount - 1
        If UCase(cmbName.List(i)) = UCase(strNewName) Then
            If MsgBox(strNewName & " already in use, overwrite?", vbYesNo) <> vbYes Then
                Exit Sub
            Else
                bDuplicate = True
            End If
            Exit For
        End If
    Next i
    ' Add to combo
    If Not bDuplicate Then cmbName.AddItem strNewName
    
    ' Save the existing toon
    Toon.SaveToon
    ' Reset Toon
    Toon.NewToon
    ' Set Name
    Toon.CharName = strNewName
    ' Load interface
    LoadFormFromChar
    
End Sub

Private Sub mnuCharacterRename_Click()
' Rename the char
Dim strNewName As String
Dim strOldName As String
Dim i As Integer
Dim bDuplicate As Boolean

    strOldName = cmbName.Text
    ' Get name
    strNewName = InputBox("Please enter a new name for " & strOldName)
    If strNewName = vbNullString Then Exit Sub
    ' Remove the old name
    For i = cmbName.ListCount - 1 To 0 Step -1
        If cmbName.List(i) = strOldName Then
            cmbName.RemoveItem i
        End If
    Next i
    ' Check for duplicates
    For i = 0 To cmbName.ListCount - 1
        If UCase(cmbName.List(i)) = UCase(strNewName) Then
            If MsgBox(strNewName & " already in use, overwrite?", vbYesNo) <> vbYes Then
                Exit Sub
            Else
                bDuplicate = True
            End If
            Exit For
        End If
    Next i
    ' Add to combo
    If Not bDuplicate Then cmbName.AddItem strNewName
    ' Set Name
    Toon.CharName = strNewName
    ' Load interface
    LoadFormFromChar
    ' Delete the old data
    DeleteCharData strOldName
    
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpHelp_Click()
    frmHelp.Show vbModal, Me
End Sub

Private Sub mnuIntenetShadows_Click()
Dim strURL As String

    strURL = "http://www.shadowbrothers.org/"
    ShellEx strURL, , , , , Me.hWnd
    
End Sub

Private Sub mnuInternetPointer_Click()
Dim strURL As String

    strURL = "http://www.memphis.co.uk/pointer/"
    ShellEx strURL, , , , , Me.hWnd
    
End Sub

Private Sub mnuInternetUpdate_Click()
    Update
End Sub

Private Sub tabGrids_Click()

    If tabGrids.SelectedItem.index = mintCurGrid Then Exit Sub  ' No need to change frame.
    ' Otherwise, hide old frame, show new.
    Select Case tabGrids.SelectedItem.index
        Case 1
            grdAttributes.Visible = True
        Case 2
            grdSpec.Visible = True
        Case 3
            grdNonSpec.Visible = True
        Case Else
    End Select
    Select Case mintCurGrid
        Case 1
            grdAttributes.Visible = False
        Case 2
            grdSpec.Visible = False
        Case 3
            grdNonSpec.Visible = False
        Case Else
    End Select
    
    ' Set Store the current grid
    mintCurGrid = tabGrids.SelectedItem.index

End Sub

Private Sub txtAttributeBuff_Change()
Dim lngValue As Long
Dim intAttribute As Integer

    ' Calculate the index of the Attribute we are editing
    intAttribute = mintCellIndex Mod 9
    
    ' Get the value
    lngValue = Val(txtAttributeBuff.Text)
    If (lngValue < 1000) And (lngValue > -1000) Then
        ' Set the toon attribute
        Toon.BuffedAttribute(intAttribute) = lngValue
        ' Set the grid text
        grdAttributes.TextArray(mintCellIndex) = lngValue
        ' Set the total for that attribute
        grdAttributes.TextArray((ATTRIBUTETOTAL * grdAttributes.Cols) + intAttribute) = Toon.TotalAttribute(intAttribute)
    Else
        MsgBox "Yeah, like a " & lngValue & " buff is going to happen!"
        SelectText txtAttributeBuff
    End If
    
End Sub

Private Sub txtAttributeBuff_GotFocus()

    SelectText txtAttributeBuff

End Sub

Private Sub txtAttributeBuff_LostFocus()

    grdAttributes.SetFocus
    txtAttributeBuff.Visible = False
    
End Sub

Private Sub txtAttributeAdditional_Change()
Dim lngValue As Long
Dim intAttribute As Integer
Dim intTotal As Integer
Dim i As Integer

    ' Get the value
    lngValue = Val(txtAttributeAdditional.Text)
    If (lngValue < 0) Or (lngValue > 30) Then Exit Sub
    
    ' Calculate the index of the Attribute we are editing
    intAttribute = mintCellIndex Mod 9
            
    With grdAttributes
        ' Calculate the total for that attribute
        For i = LBound(AttributeArray) To UBound(AttributeArray)
            intTotal = intTotal + Toon.AdditionalAttributePoints(i)
        Next i
        ' Check we have no more than 30 points
        If intTotal - Val(.TextArray(mintCellIndex)) + lngValue <= 30 Then
            ' Set the toon attribute
            Toon.AdditionalAttributePoints(intAttribute) = lngValue
            ' Set the grid text
            grdAttributes.TextArray(mintCellIndex) = Toon.AdditionalAttributePoints(intAttribute)
            ' Set the total for that attribute
            grdAttributes.TextArray((ATTRIBUTETOTAL * .Cols) + intAttribute) = Toon.TotalAttribute(intAttribute)
        Else
            MsgBox "Sorry, maximum spend of 30 points."
            SelectText txtAttributeAdditional
        End If
    End With
    
End Sub

Private Sub txtAttributeAdditional_GotFocus()

    SelectText txtAttributeAdditional

End Sub

Private Sub txtAttributeAdditional_LostFocus()

    grdAttributes.SetFocus
    txtAttributeAdditional.Visible = False
    
End Sub

Private Sub txtDescription_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Launch further description if there is one

    If Button = vbRightButton Then
        If False Then
            ' Shell to URL
        Else
            MsgBox "Sorry, no further information is available currently"
        End If
    End If
    
End Sub

Private Sub txtLevel_GotFocus()

    txtLevel.SelStart = 0
    txtLevel.SelLength = Len(txtLevel.Text)
    
End Sub

Private Sub txtLevel_LostFocus()
    txtLevel.Text = Toon.Level
End Sub

Private Sub txtSpec_Change()
Dim intPointsRequired As Integer
Dim intLine As Integer
    
    ' Calculate line
    intLine = mintCellIndex Mod grdSpec.Cols
    If (Val(txtSpec.Text) < 10000) And (Val(txtSpec.Text) >= 0) Then
        ' are we changing based on percent, level, or points
        Select Case Int(mintCellIndex / grdSpec.Cols)
            Case SPECPERCENT
                Toon.LinePercent(intLine) = Val(txtSpec.Text)
            Case SPECPOINTS
                Toon.LinePoints(intLine) = Val(txtSpec.Text)
            Case SPECLEVEL
                Toon.LineLevel(intLine) = Val(txtSpec.Text)
            Case Else
        End Select
        
        With grdSpec
            .TextArray((SPECPERCENT * .Cols) + intLine) = Toon.LinePercent(intLine)
            .TextArray((SPECPOINTS * .Cols) + intLine) = Toon.LinePoints(intLine)
            .TextArray((SPECLEVEL * .Cols) + intLine) = Toon.LineLevel(intLine)
        End With
    
        SetSpecPoints
        SetAvailableSpecStyles intLine
    Else
        MsgBox "Please enter a slightly more realistic value!"
        SelectText txtSpec
    End If
    
End Sub

Private Sub txtSpec_GotFocus()

    txtSpec.SelStart = 0
    txtSpec.SelLength = Len(txtSpec.Text)

End Sub

Private Sub txtSpec_LostFocus()

    grdSpec.SetFocus
    txtSpec.Visible = False
    
End Sub

Private Sub txtLevel_Change()
Dim dblLevel As Double
Dim dblLevelOld As Double

    If Val(txtLevel.Text) = 0 Then
        dblLevel = 0
    Else
        dblLevel = CDbl(txtLevel.Text)
    End If
    If dblLevel = Toon.Level Then Exit Sub
    
    If dblLevel > MAXLEVEL Then
        MsgBox "Please enter a level between 5 and " & MAXLEVEL
        SelectText txtLevel
    Else
        dblLevelOld = Toon.Level
        Toon.Level = dblLevel
        If Toon.PointsAvailable < 0 Then
        Toon.Level = dblLevelOld
'        MsgBox "Level to low for spec points spend"
'        SelectText txtLevel
        Else
            SetEarnedAttributes
            txtPointsAcquired.Text = Toon.PointsAcquired
            Toon.RecalculateSpecLines
            SetSpecPointsSpend
            SetAvailableSpecStyles
            SetAvailableNonSpecStyles
            SetAvailableArmor
            SetAvailableOther
        End If
    End If
    
End Sub

Private Sub txtPointsAcquired_Change()
    txtPointsAvailable.Text = Toon.PointsAvailable

'    txtPointsAvailable.Text = Val(txtPointsAcquired.Text) - Val(txtPointsSpent.Text)
End Sub

Private Sub txtPointsSpent_Change()
    txtPointsAvailable.Text = Toon.PointsAvailable
'        txtPointsAvailable.Text = Val(txtPointsAcquired.Text) - Val(txtPointsSpent.Text)
End Sub
