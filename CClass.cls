VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'properties
Private pClassID As Long
Private pCourse As String
Private pSection As String
Private pBlock As Long
Private pFaculty As String

'ID property Getter
Public Property Get ClassID() As Long
    ClassID = pClassID
End Property
'ID property Setter
Public Property Let ClassID(value As Long)
    pClassID = value
End Property

'property Getter
Public Property Get Course() As String
    Course = pCourse
End Property

'property Setter
Public Property Let Course(value As String)
    pCourse = value
End Property
'property Getter
Public Property Get Section() As Long
    Section = pSection
End Property

'property Setter
Public Property Let Section(value As Long)
    pSection = value
End Property
'property Getter
Public Property Get Block() As Long
    Block = pBlock
End Property

'property Setter
Public Property Let Block(value As Long)
    pBlock = value
End Property
      
'property Getter
Public Property Get faculty() As String
    faculty = pFaculty
End Property

'property Setter
Public Property Let faculty(value As String)
    pFaculty = value
End Property
