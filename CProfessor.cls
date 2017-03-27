VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CProfessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pProfessorID As Long
Private pProfessorName As String
Private pProfessorType As String
Private pTerminalDegree As Long
Private pPreferredCourse As String
Private pPreferredBlocks As Collection

Public Property Get ProfessorID() As Long
    ProfessorID = pProfessorID
End Property

Public Property Let ProfessorID(value As Long)
    pProfessorID = value
End Property
Public Property Get ProfessorName() As String
    ProfessorName = pProfessorName
End Property

Public Property Let ProfessorName(value As String)
    pProfessorName = value
End Property
'Type property Getter
Public Property Get ProfessorType() As String
    ProfessorType = pProfessorType
End Property

'Type property Setter
Public Property Let ProfessorType(value As String)
    pProfessorType = value
End Property
Public Property Get TerminalDegree() As Long
    TerminalDegree = pTerminalDegree
End Property

Public Property Let TerminalDegree(value As Long)
    pTerminalDegree = value
End Property
Public Property Get PreferredCourse() As String
    PreferredCourse = pPreferredCourse
End Property

Public Property Let PreferredCourse(value As String)
    pPreferredCourse = value
End Property
Public Property Get preferredBlocks() As Collection
    Set preferredBlocks = pPreferredBlocks
End Property

Public Property Let preferredBlocks(value As Collection)
    Set pPreferredBlocks = value
End Property

