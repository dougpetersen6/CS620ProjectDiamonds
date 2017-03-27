VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPreferredBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pPreferredBlockID As Long
Private pProfessorName As String
Private pPreferredLevel As Long

Public Property Get PreferredBlockID() As Long
    PreferredBlockID = pPreferredBlockID
End Property

Public Property Let PreferredBlockID(value As Long)
    pPreferredBlockID = value
End Property
Public Property Get ProfessorName() As String
    ProfessorName = pProfessorName
End Property

Public Property Let ProfessorName(value As String)
    pProfessorName = value
End Property

Public Property Get PreferredLevel() As Long
    PreferredLevel = pPreferredLevel
End Property

Public Property Let PreferredLevel(value As Long)
    pPreferredLevel = value
End Property
