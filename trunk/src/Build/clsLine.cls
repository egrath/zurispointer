VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsStyle"
Attribute VB_Ext_KEY = "Member0" ,"clsStyle"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection
Private mvarLineName As String
Private mvarDescription As String
Private mvarLevel As Integer
Private mvarCap As Integer

Public Property Let LineName(ByVal vData As String)
    mvarLineName = vData
End Property

Public Property Get LineName() As String
    LineName = mvarLineName
End Property

Public Property Let Description(ByVal vData As String)
    mvarDescription = vData
End Property

Public Property Get Description() As String
    Description = mvarDescription
End Property

Public Property Let Level(ByVal vData As Integer)
    mvarLevel = vData
End Property

Public Property Get Level() As Integer
    Level = mvarLevel
End Property

Public Property Let Cap(ByVal vData As Integer)
    mvarCap = vData
End Property

Public Property Get Cap() As Integer
    Cap = mvarCap
End Property

Public Function Add(StyleName As String, Level As Integer, Description As String, Optional sKey As String) As clsStyle
    'create a new object
    Dim objNewMember As clsStyle
    Set objNewMember = New clsStyle

    'set the properties passed into the method
    objNewMember.StyleName = StyleName
    objNewMember.Level = Level
    objNewMember.Description = Description
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As clsStyle
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

