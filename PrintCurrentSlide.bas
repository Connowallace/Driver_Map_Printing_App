Attribute VB_Name = "Module1"
' Function for printing the slide currently shown on the screen
' PLEASE NOTE: This function only works during a slide show
Sub PrintMe()

    Dim lCurrentSlide As Long

    ' Get the SlideID of the slide currently in view
    lCurrentSlide = SlideShowWindows(1).View.Slide.SlideNumber

    ' Set up print options
    With ActivePresentation.PrintOptions

        ' Print a range that includes only the current slide
        .RangeType = ppPrintSlideRange
        ' Change it to .RangeType = ppPrintAll to print the entire presentation
        ' You may also need to delete the following four lines to print all
        With .Ranges
            .ClearAll
            .Add Start:=lCurrentSlide, End:=lCurrentSlide
        End With

        .NumberOfCopies = 1

         ' This prints notes pages;  change it to e.g. ppPrintOutputSlides to print slides
         ' To see the other types delete everything from the = sign to the end of the line below
         ' Then type = at the end of the line;  VBA's Intellisense feature will show you the available options
        .OutputType = ppPrintOutputSlides

        .PrintHiddenSlides = msoTrue

         ' Likewise, change this if you want color or pure b/w
        .PrintColorType = ppPrintBlackAndWhite

        .FitToPage = msoFalse
        .FrameSlides = msoFalse

    End With

    ' and PRINT
    ActivePresentation.PrintOut
    Debug.Print "print done"

End Sub
