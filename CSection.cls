VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pSectionID As Long
Private pSectionType As String
Public Property Get SectionID() As Long
    SectionID = pSectionID
End Property

Public Property Let SectionID(value As Long)
    pSectionID = value
End Property

Public Property Get SectionType() As String
    SectionType = pSectionType
End Property

Public Property Let SectionType(value As String)
    pSectionType = value
End Property
Sub SectionOverBooked()
    MsgBox "Cannot assign sections more than 4 to a professor"
End Sub

