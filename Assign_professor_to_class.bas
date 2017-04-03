Attribute VB_Name = "Assign_professor_to_class"
'This module reads class schedule on the "section list" tab and generate a list of professors eligible to assign to a class

Sub Assign_professor_to_class()

Worksheets("Sections List").Activate
Range("A2").Activate
nmbrClass = ActiveCell.Offset(-1, 4).value


'This module reads professors' info(full/part-time & has (no) terminal degree) and preferred blocks and courses

Dim professor As CProfessor
Dim preferredBlock As CPreferredBlock
Dim professors As Collection
Dim nmbrPro As Long
Dim proCntr As Long
Dim nmbrBlock As Long

Dim nmbrCourse As Long
Dim preferredCourse As CPreferredCourse
nmbrCourse = ActiveSheet.Cells(1, Columns.count).End(xlToLeft).Column - 38 '38 is column number from the left most to the last block

Set professors = New Collection

'read how many professors to assign
nmbrPro = ActiveCell.Offset(0, 5).value


'add professor to professors and assign preferredBlocks and preferredCourses
For proCntr = 1 To nmbrPro
    
    Set professor = New CProfessor
    'add professor to collection professors
    professors.Add professor
    
    'instantiate a preferred blocks collection
    Dim prefBlocks As Collection
    Set prefBlocks = New Collection
    'instantiate a preferred courses collection
    Dim prefCourses As Collection
    Set prefCrouses = New Collection
    
    
    professor.ProfessorName = ActiveCell.Offset(proCntr - 1, 6).value
    professor.ProfessorID = proCntr
    professor.ProfessorType = ActiveCell.Offset(proCntr - 1, 7).value
    professor.TerminalDegree = ActiveCell.Offset(proCntr - 1, 8).value
    
    'read preferred blocks
    For nmbrBlock = 1 To 28
        Set preferredBlock = New CPreferredBlock
        If ActiveCell.Offset(proCntr - 1, (9 + nmbrBlock)).value < 10 Then
        preferredBlock.PreferredBlockID = nmbrBlock
        preferredBlock.PreferredLevel = ActiveCell.Offset(proCntr - 1, (9 + nmbrBlock)).value 'from col A to col K is 9 cols in between
        preferredBlock.ProfessorName = professor.ProfessorName
        End If
        prefBlocks.Add preferredBlock
    Next
    'link preferredblocks to the professor. It's optional since every preferredblock has a professor
    professor.preferredBlocks = prefBlocks
    
    'read preferred courses
    For courseCntr = 1 To nmbrCourse
        Set preferredCourse = New CPreferredCourse
        If ActiveCell.Offset(proCntr - 1, (37 + courseCntr)).value < 10 Then
        preferredCourse.PreferredCourseID = courseCntr
        preferredCourse.PreferredCourseName = ActiveCell.Offset(-1, 37 + courseCntr).value
        preferredCourse.PreferredLevel = ActiveCell.Offset(proCntr - 1, (37 + courseCntr)).value 'from col A to col AM is 37 cols in between
        preferredCourse.ProfessorName = professor.ProfessorName
        End If
        prefCrouses.Add preferredCourse
    Next
    'link preferredblocks to the professor. It's optional since every preferredblock has a professor
    professor.preferredCourses = prefCrouses
Next


'initiate a classes collection
Dim classes As Collection
Set classes = New Collection

'read data into the class
For classCntr = 1 To nmbrClass
    Dim Class As CClass
    Set Class = New CClass
    Class.ClassID = classCntr
    Class.Course = Left(ActiveCell.Offset(classCntr - 1, 0).value, 6)
    Class.Section = Right(ActiveCell.Offset(classCntr - 1, 0).value, 3)
    Class.blockID = ActiveCell.Offset(classCntr - 1, 1).value
    
    classes.Add Class
Next
    
'generate a list of professors eligible to assign to a class
For Each Class In classes
    For p = 1 To professors.count
            For N = 1 To professors(p).preferredBlocks.count
                If Class.blockID = professors(p).preferredBlocks.Item(N).PreferredBlockID Then
                    For m = 1 To professors(p).preferredCourses.count
                        If Class.Course = professors(p).preferredCourses.Item(m).PreferredCourseName Then
                        'Debug.Print professors(p).ProfessorName
                        Class.AddToList
                        professors(p).AddToList
                        professors(p).preferredCourses.Item(m).AddToList
                        professors(p).preferredBlocks.Item(N).AddToList
                        End If
                    Next m
                    
                End If
            Next N
    
    Next p
Next

'Need assignment method, say I have four professors eligible for course CS150 at block 1. Course CS150 needs 3 sections and thus 3 professors.
'Find the best assignment so that professors with 0 in block 1 and course CS150 will teach the course


                    
'display the professor-class assignment
'Range("A2").Activate
'For classCntr = 1 To nmbrClass
    'ActiveCell.Offset(classCntr - 1, 2).value = classes(classCntr).Course & "-" & CStr(classes(classCntr).Section) & classes(classCntr).faculty
'Next

Range("A500").Select
End Sub
