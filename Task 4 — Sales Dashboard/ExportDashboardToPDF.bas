Sub ExportDashboardToPDF()
    Dim path As String
    path = ThisWorkbook.path & "\Dashboard_Export_" & Format(Date, "yyyy-mm-dd") & ".pdf"
    
    With Sheets("Dashboard")
        .PageSetup.Orientation = xlLandscape
        .PageSetup.Zoom = False
        .PageSetup.FitToPagesWide = 1
        .PageSetup.FitToPagesTall = 99
        .PageSetup.PaperSize = xlPaperA4
        .PageSetup.LeftMargin = Application.InchesToPoints(0.4)
        .PageSetup.RightMargin = Application.InchesToPoints(0.4)
        .PageSetup.TopMargin = Application.InchesToPoints(0.5)
        .PageSetup.BottomMargin = Application.InchesToPoints(0.5)
        .PageSetup.PrintArea = .Range("A1:K41").Address
        
        .ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=path, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=True
    End With
End Sub
