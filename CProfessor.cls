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
Private pPreferredCourses As Collection 'not sure if they should be array of strings
Private pPreferredBlocks As Collection 'not sure if they should be array of Longs

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
Public Property Get preferredCourses() As Collection
   Set preferredCourses = pPreferredCourses
End Property

Public Property Let preferredCourses(value As Collection)
   Set pPreferredCourses = value
End Property
Public Property Get preferredBlocks() As Collection
    Set preferredBlocks = pPreferredBlocks
End Property

Public Property Let preferredBlocks(value As Collection)
    Set pPreferredBlocks = value
End Property

Public Sub AddToList()
    Worksheets("Sections List").Select
    Range("B499").End(xlDown).Offset(1, 0).Select
    'add id, name, type, degree
    ActiveCell.value = Me.ProfessorID
    ActiveCell.Offset(0, 1).value = Me.ProfessorName
    ActiveCell.Offset(0, 2).value = Me.ProfessorType
    ActiveCell.Offset(0, 3).value = Me.TerminalDegree
    
End Sub
