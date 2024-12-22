Dim pptSlide As Slide
Dim pptShape As Shape
Dim pptTable As Table

Dim CNFont As String
Dim EnFont As String

Sub main()
    ' Move your cursor to here, and click F5, it will change your font in all pages including shapes, tables, but font in smart art shape are excluded.
    Call SetShapeFont
End Sub

Sub SetShapeFont()
    Dim i As Long
    Dim J As Long
    Dim txt As String
    
    CNFont = "Microsoft YaHei" ' Font for Chinese
    EnFont = "Microsoft YaHei" ' Font for English
    
    With ActivePresentation
        For Each pptSlide In .Slides
            For Each pptShape In pptSlide.Shapes
                With pptShape
                    If .HasTextFrame Then
                        If .TextFrame.HasText Then
                            ' Retrieve the current text
                            txt = .TextFrame.TextRange.Text
                            
                            ' Replace every space with two spaces
                            txt = Replace(txt, "          ", "               ") ' 10 spaces to 15 spaces
                            
                            ' Apply modified text
                            .TextFrame.TextRange.Text = txt
                            
                            ' Font for Chinese and Japanese
                            .TextFrame.TextRange.Font.NameFarEast = CNFont
                            ' Font for English
                            .TextFrame.TextRange.Font.Name = EnFont
                        End If
                    End If
                End With
                
                ' Handle tables
                If pptShape.HasTable Then
                    Set pptTable = pptShape.Table
                    For i = 1 To pptTable.Columns.Count
                        For J = 1 To pptTable.Rows.Count
                            ' Retrieve and modify table cell text
                            txt = pptTable.Cell(J, i).Shape.TextFrame.TextRange.Text
                            txt = Replace(txt, "          ", "                    ") ' Replace 10 spaces with 20
                            
                            ' Apply modified text
                            pptTable.Cell(J, i).Shape.TextFrame.TextRange.Text = txt
                            
                            ' Font for Chinese and Japanese
                            pptTable.Cell(J, i).Shape.TextFrame.TextRange.Font.NameFarEast = CNFont
                            ' Font for English
                            pptTable.Cell(J, i).Shape.TextFrame.TextRange.Font.Name = EnFont
                        Next J
                    Next i
                End If
            Next
        Next
    End With
End Sub
