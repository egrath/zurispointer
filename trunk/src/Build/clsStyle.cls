VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarStyleName As String
Private mvarLevel As Integer
Private mvarDescription As String

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

Public Property Let StyleName(ByVal vData As String)
    mvarStyleName = vData
End Property

Public Property Get StyleName() As String
    StyleName = mvarStyleName
End Property



