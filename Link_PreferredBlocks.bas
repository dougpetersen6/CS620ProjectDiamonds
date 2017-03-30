Attribute VB_Name = "Link_PreferredBlocks"
'This module reads professors' info(full/part-time & has (no) terminal degree) and preferred blocks
Sub Link_PreferredBlocks()

Dim professor As CProfessor
Dim preferredBlock As CPreferredBlock
Dim professors As Collection
Dim nmbrPro As Long
Dim proCntr As Long
Dim nmbrBlock As Long
Dim nmbrCourse As Long
Dim preferredCourse As CPreferredCourse

Set professors = New Collection

'read how many professors
nmbrPro = Worksheets("Block Preference").Range("A2").value
Worksheets("Block Preference").Activate
Range("B2").Activate
        
'count number of courses
nmbrCourse = Sheet1.Cells(1, Columns.Count).End(xlToLeft).Column


'add professor to professors and assign preferredBlocks
For proCntr = 0 To nmbrPro
    
    Set professor = New CProfessor
    'add professor to collection professors
    professors.Add professor
    
    'instantiate a preferred blocks collection
    Dim prefBlocks As Collection
    Set prefBlocks = New Collection
   'instantiate a preferred courses collection
    Dim prefCourses As Collection
    Set prefCrouses = New Collection
    
    professor.ProfessorName = ActiveCell.Offset(proCntr, 0).value
    professor.ProfessorID = proCntr + 1
    professor.ProfessorType = ActiveCell.Offset(proCntr, 1).value
    professor.TerminalDegree = ActiveCell.Offset(proCntr, 2).value
    
    'read preferred blocks
    For nmbrBlock = 1 To 28
        Set preferredBlock = New CPreferredBlock

        
        preferredBlock.PreferredBlockID = nmbrBlock
        preferredBlock.PreferredLevel = ActiveCell.Offset(proCntr, (2 + nmbrBlock)).value
        preferredBlock.ProfessorName = professor.ProfessorName
        
        prefBlocks.Add preferredBlock
     'preferred course
                
    'read preferred blocks
    For nmbrCourse = 1 To numbrCourse
        Set preferredCourse = New CPreferredCourse

        
        preferredCourse.PreferredBlockID = nmbrCourse
        preferredCourse.PreferredLevel = ActiveCell.Offset(proCntr, (30 + nmbrBlock)).Value
        preferredCourse.ProfessorName = professor.ProfessorName
        
        prefCourse.Add preferredCourse
    Next
    'link preferredblocks to the professor. It's optional since every preferredblock has a professor
    professor.preferredBlocks = prefBlocks

  
    

Next
'test
Debug.Print professors(20).ProfessorName
Debug.Print professors(20).preferredBlocks(27).PreferredBlockID
Debug.Print professors(20).preferredBlocks(27).PreferredLevel

End Sub
