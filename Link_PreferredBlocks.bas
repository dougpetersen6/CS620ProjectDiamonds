Attribute VB_Name = "Link_PreferredBlocks"
'This module reads professors' info(full/part-time & has (no) terminal degree) and preferred blocks
Sub Link_PreferredBlocks()

Dim professor As CProfessor
Dim preferredBlock As CPreferredBlock
Dim professors As Collection
Dim nmbrPro As Long
Dim proCntr As Long
Dim nmbrBlock As Long


Set professors = New Collection

'read how many professors
nmbrPro = Worksheets("Sections List").Range("F2").value
Worksheets("Sections List").Activate
Range("G2").Activate

'add professor to professors and assign preferredBlocks
For proCntr = 0 To nmbrPro
    
    Set professor = New CProfessor
    'add professor to collection professors
    professors.Add professor
    
    'instantiate a preferred blocks collection
    Dim prefBlocks As Collection
    Set prefBlocks = New Collection
    
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
    Next
    'link preferredblocks to the professor. It's optional since every preferredblock has a professor
    professor.preferredBlocks = prefBlocks

Next
'test
Debug.Print professors(20).ProfessorName
Debug.Print professors(20).preferredBlocks(27).PreferredBlockID
Debug.Print professors(20).preferredBlocks(27).PreferredLevel

End Sub
        
'Return a list of professors who prefer to teach at a certain blockID
Function findProfByBlock(blockID As Long) As String
'define a collection to store all professors
Dim professors As Collection
Set professors = New Collection
'Initiate return value to be an empty string
findProfByBlock = ""

Dim proCntr As Long 'define a counter
Dim nmbrPro As Long  'number of total professors
nmbrPro = Worksheets("Sections List").Range("F2").value

Dim block As Range
Worksheets("Sections List").Activate
'compare each block with the passed in blockID
For Each block In [Blocks]
    If block.value = blockID Then 'find the matching block
     
     For proCntr = 1 To nmbrPro
        If (Worksheets("Sections List").Cells(proCntr + 1, block.Column).value = 0) Then 'find the preferedScale =0
        professors.Add (Worksheets("Sections List").Cells(proCntr + 1, 7)) 'add the corresponding professor into the collection
        End If
     Next proCntr
     Exit For 'exit for loop since there is only one block to be the same as the passed in blockID
     End If
Next block

'loop the professors collection
For p = 1 To professors.Count
findProfByBlock = findProfByBlock & professors.Item(p) & " . " 'assign collection items to be returned
Next p

End Function
