Private Sub CommandButton1_Click()

Dim rng As Range
Dim PowerPointApp As Object
Dim myPresentation As Object
Dim mySlide As Object
Dim myShape As Object

    Dim wks As Worksheet
    Set wks = ActiveSheet
    Dim rowRange As Range
    Dim colRange As Range
    Dim LastCol As Long
    Dim LastRow As Long
    Dim i, slideCtr, j As Integer
    Dim PPT As Object
    Dim objTextBox As Object 
    Dim objSlide As Object
	
	

        LastRow = wks.Cells(Rows.Count, 1).End(xlUp).Row
        Set rowRange = wks.Range("A1:A" & LastRow)
        Set PPT = CreateObject("PowerPoint.Application")
        PPT.Visible = True
        Set myPresentation = PPT.Presentations.Open("C:\Users\melme\Desktop\Presentation.pptx")
        slideCtr = 1
        
  'Optimize Code
    Application.ScreenUpdating = False
  
    For Each Row In rowRange
        i = i + 1
          
        Set mySlide = myPresentation.Slides(i).Duplicate
        LastCol = Cells(i, Columns.Count).End(xlToLeft).Column
        Set colRange = wks.Range(wks.Cells(i, 1), wks.Cells(i, LastCol))
                  
              Set objSlide = myPresentation.Slides.Item(i)
            
            
        For Each cell In colRange
        j = j + 1
            cell.Copy
            'Set myShape = mySlide.Shapes(mySlide.Shapes.Count)
         
            Set objTextBox = objSlide.Shapes.Item(j)

            objTextBox.TextFrame.TextRange.Text = cell.Value
            
            'Make PowerPoint Visible and Active
              PPT.Visible = True
              PPT.Activate
            
            'Clear The Clipboard
              Application.CutCopyMode = False
        Next cell
		j = 0
    Next Row

End Sub
