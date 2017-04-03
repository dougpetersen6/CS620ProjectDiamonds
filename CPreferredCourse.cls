VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPreferredCourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pPreferredCourseID As Long
Private pPreferredCourseName As String
Private pProfessorName As String
Private pPreferredLevel As Long

Public Property Get PreferredCourseID() As Long
    PreferredCourseID = pPreferredCourseID
End Property

Public Property Let PreferredCourseID(value As Long)
    pPreferredCourseID = value
End Property
Public Property Get PreferredCourseName() As String
    PreferredCourseName = pPreferredCourseName
End Property

Public Property Let PreferredCourseName(value As String)
    pPreferredCourseName = value
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
Public Sub AddToList()
    Worksheets("Sections List").Select
    Range("F499").End(xlDown).Offset(1, 0).Select
    'add preferred level
    ActiveCell.value = Me.PreferredLevel
    
End Sub
