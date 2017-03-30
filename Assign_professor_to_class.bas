'This is Doug's copy of Master branch

Attribute VB_Name = "Assign_professor_to_class"
Sub Assign_professor_to_class()

Worksheets("Sections List").Activate
Range("A2").Activate
nmbrClass = ActiveCell.Offset(-1, 4).value

Dim classes As Collection
Set classes = New Collection

'read data into the class
For classCntr = 1 To nmbrClass
    Dim Class As CClass
    Set Class = New CClass
    Class.ClassID = classCntr
    Class.Course = Left(ActiveCell.Offset(classCntr - 1, 0).value, 5)
    Class.Section = Right(ActiveCell.Offset(classCntr - 1, 0).value, 3)
    Class.Block = ActiveCell.Offset(classCntr - 1, 1).value
    
    classes.Add Class
Next
    
'assign a class to a random professor
For Each Class In classes
    For Each p In [professors]
        If Not IsEmpty(p) Then
        classes(Int(nmbrClass * Rnd()) + 1).faculty = p
        'Debug.Print classes(Int(nmbrClass * Rnd()) + 1).ClassID
        
        End If
    Next
Next

'display the professor assigned to a class
Range("A2").Activate
For classCntr = 1 To nmbrClass
    ActiveCell.Offset(classCntr - 1, 2).value = classes(classCntr).faculty
Next

'need totalScaleOfCourse (classes)
'need totalScaleOfBlock (classes)

'test
Debug.Print classes(10).faculty

End Sub
