VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'properties
Private pCourseID As Long
Private pCourseName As String
Private pCourseType As String

'ID property Getter
Public Property Get CourseID() As Long
    CourseID = pCourseID
End Property
'ID property Setter
Public Property Let CourseID(value As Long)
    pCourseID = value
End Property

'Name property Getter
Public Property Get CourseName() As String
    CourseName = pCourseName
End Property

'Name property Setter
Public Property Let CourseName(value As String)
    pCourseName = value
End Property
'Type property Getter
Public Property Get CourseType() As String
    CourseType = pCourseType
End Property

'Type property Setter
Public Property Let CourseType(value As String)
    pCourseType = value
End Property
