'VBA Create Script
    'Create a script that will loop through one year of stock data for each run and return the total volume each stock had over that year.
    'You will also need to display the ticker symbol to coincide with the total stock volume.
    'Your result should look as follows (note: all solution images are for 2015 data).
    
Sub LoopThroughTotalColumn()
    For Each ws In Worksheets
                
        'Declare each variable for each TickerElement
        
        'A
        Dim CountTotalTickerA As Long
        CountTotalTickerA = 0
        Dim OpenCloseArrayA() As String
        Dim DailyAverageDifferenceA As Double
        Dim FinalDailyAverageA As Double
        Dim YearlyChangeA As Double
        Dim PercentChangeA As Double
        Dim VolumeArrayA() As String
        Dim TotalStockVolumeA As Long
        
        
        'AA
        Dim CountTotalTickerAA As Long
        CountTotalTickerAA = 0
        Dim OpenCloseArrayAA() As String
        Dim DailyAverageDifferenceAA As Double
        Dim FinalDailyAverageAA As Double
        Dim YearlyChangeAA As Double
        Dim PercentChangeAA As Double
        Dim VolumeArrayAA() As String
        Dim TotalStockVolumeAA As Long
        
        
        'AA-B
        Dim CountTotalTickerAAB As Long
        CountTotalTickerAAB = 0
        Dim OpenCloseArrayAAB() As String
        Dim DailyAverageDifferenceAAB As Double
        Dim FinalDailyAverageAAB As Double
        Dim YearlyChangeAAB As Double
        Dim PercentChangeAAB As Double
        Dim VolumeArrayAAB() As String
        Dim TotalStockVolumeAAB As Long
        
        
        'AAC
        Dim CountTotalTickerAAC As Long
        CountTotalTickerAAC = 0
        Dim OpenCloseArrayAAC() As String
        Dim FinalDailyAverageAAC As Double
        Dim DailyAverageDifferenceAAC As Double
        Dim YearlyChangeAAC As Double
        Dim PercentChangeAAC As Double
        Dim VolumeArrayAAC() As String
        Dim TotalStockVolumeAAC As Long
        
        
        'AAN
        Dim CountTotalTickerAAN As Long
        CountTotalTickerAAN = 0
        Dim OpenCloseArrayAAN() As String
        Dim FinalDailyAverageAAN As Double
        Dim DailyAverageDifferenceAAN As Double
        Dim YearlyChangeAAN As Double
        Dim PercentChangeAAN As Double
        Dim VolumeArrayAAN() As String
        Dim TotalStockVolumeAAN As Long
        
        
        'AAP
        Dim CountTotalTickerAAP As Long
        CountTotalTickerAAP = 0
        Dim OpenCloseArrayAAP() As String
        Dim FinalDailyAverageAAP As Double
        Dim DailyAverageDifferenceAAP As Double
        Dim YearlyChangeAAP As Double
        Dim PercentChangeAAP As Double
        Dim VolumeArrayAAP() As String
        Dim TotalStockVolumeAAP As Long
        
        
        'AAT
        Dim CountTotalTickerAAT As Long
        CountTotalTickerAAT = 0
        Dim OpenCloseArrayAAT() As String
        Dim FinalDailyAverageAAT As Double
        Dim DailyAverageDifferenceAAT As Double
        Dim YearlyChangeAAT As Double
        Dim PercentChangeAAT As Double
        Dim VolumeArrayAAT() As String
        Dim TotalStockVolumeAAT As Long
        
        
        'AAV
        Dim CountTotalTickerAAV As Long
        CountTotalTickerAAV = 0
        Dim OpenCloseArrayAAV() As String
        Dim FinalDailyAverageAAV As Double
        Dim DailyAverageDifferenceAAV As Double
        Dim YearlyChangeAAV As Double
        Dim PercentChangeAAV As Double
        Dim VolumeArrayAAV() As String
        Dim TotalStockVolumeAZV As Long
        
        
        'AB
        Dim CountTotalTickerAB As Long
        CountTotalTickerAB = 0
        Dim OpenCloseArrayAB() As String
        Dim FinalDailyAverageAB As Double
        Dim DailyAverageDifferenceAB As Double
        Dim YearlyChangeAB As Double
        Dim PercentChangeAB As Double
        Dim VolumeArrayAB() As String
        Dim TotalStockVolumeAB As Long
        
        
        'ABB
        Dim CountTotalTickerABB As Long
        CountTotalTickerABB = 0
        Dim OpenCloseArrayABB() As String
        Dim FinalDailyAverageABB As Double
        Dim DailyAverageDifferenceABB As Double
        Dim YearlyChangeABB As Double
        Dim PercentChangeABB As Double
        Dim VolumeArrayABB() As String
        Dim TotalStockVolumeABB As Long
        
        
        'ABBV
        Dim CountTotalTickerABBV As Long
        CountTotalTickerABBV = 0
        Dim OpenCloseArrayABBV() As String
        Dim FinalDailyAverageABBV As Double
        Dim DailyAverageDifferenceABBV As Double
        Dim YearlyChangeABBV As Double
        Dim PercentChangeABBV As Double
        Dim VolumeArrayABBV() As String
        Dim TotalStockVolumeABBV As Long
        
        
        'ABC
        Dim CountTotalTickerABC As Long
        CountTotalTickerABC = 0
        Dim OpenCloseArrayABC() As String
        Dim FinalDailyAverageABC As Double
        Dim DailyAverageDifferenceABC As Double
        Dim YearlyChangeABC As Double
        Dim PercentChangeABC As Double
        Dim VolumeArrayABC() As String
        Dim TotalStockVolumeABC As Long
        
        Dim GreatestTotalVolume() As String
        Dim GreatestTotalVolumeE() As String
        
        
        
        '=================================================================================
        'MAIN FOR LOOP INSERT KEY FUNCTIONS HERE
        '=================================================================================
        ' Collet one Columns work of data in a Array and as a total
        
        Dim ColumnCellTotal As Long
        ColumnCellTotal = (Rows.Count - 1)
        
        Dim LastCellInColumn
        ' LastCellInColumnA1 aka End of Loop
        LastCellInColumn = ws.Cells(ColumnCellTotal, 1)
        
        Dim CurrentCell As String
        
        For i = 1 To ColumnCellTotal
        
            CurrentCell = ws.Cells(i, 1).Value
            
            If CurrentCell = "String" Then
                
				
                'A
                '=================================================
                Dim ElementValue As String
                ElementValue = ws.Cells(i, 1).Value
                
                
                If (StrComp(ElementValue, "A", vbTextCompare) = 0) Then
                
                    'Grab Open
                    Dim OpenValue As Double
                    OpenValue = Cells(3, i).Value
                    
                    'Grab close
                    Dim CloseValue As Double
                    CloseValue = Cells(3, i).Value
                    
                    'Function calculate difference(opne, close) for current date
                    Dim OpenCloseDifferenceA As Double
                    OpenCloseDifferenceA = (OpenValue - CloseValue)
                    
                    CountTotalTickerA = CountTotalTickerA + 1
                    OpenCloseArrayA [i] = OpenCloseDifferenceA
                    
                    Dim RowVolumeA As Long
                    RowVolumeA = Cells(7, i).Value
                    VolumeArrayA.Add(  [i] = RowVolumeA
                    
                End If
                
                'Opps Made this by accadent Calculate Daily Average Change
                '=================================================
                For Each element In OpenCloseArrayA
                    DailyAverageDifferenceA = (DailyAverageDifferenceA + CDbl(element))
                FinalDailyAverageA = (DailyAverageDifferenceA / CountTotalTickerA)
                
                
                
                'AA
                '=================================================
                If (StrComp(ElementValue, "AA", vbTextCompare) = 0) Then
                
                
                    'Grab Open
                    Dim OpenValue As Double
                    OpenValue = Cells(3, i).Value
                    
                    'Grab close
                    Dim CloseValue As Double
                    CloseValue = Cells(3, i).Value
                    
                    'Function calculate difference(opne, close) for current date
                    Dim OpenCloseDifferenceAA As Double
                    OpenCloseDifferenceAA = (OpenValue - CloseValue)
                    
                    CountTotalTickerAA = CountTotalTickerAA + 1
                    OpenCloseArrayAA [i] = OpenCloseDifferenceAA
                    
                    Dim RowVolumeAA As Long
                    RowVolumeAA = Cells(7, i).Value
                    VolumeArrayAA [i] = RowVolumeAA
                    
                End If
                
                For Each element In OpenCloseArrayAA
                    DailyAverageDifferenceAA = (DailyAverageDifferenceAA + CDbl(element))
                
                FinalDailyAverageAA = (DailyAverageDifferenceAA / CountTotalTickerAA)
                
                
                'AA-B
                '=================================================
                
                If (StrComp(ElementValue, "AA-B", vbTextCompare) = 0) Then
                
                
                    'Grab Open
                    Dim OpenValue As Double
                    OpenValue = Cells(3, i).Value
                    
                    'Grab close
                    Dim CloseValue As Double
                    CloseValue = Cells(3, i).Value
                    
                    'Function calculate difference(opne, close) for current date
                    Dim OpenCloseDifferenceAAB As Double
                    OpenCloseDifferenceAAB = (OpenValue - CloseValue)
                    
                    CountTotalTickerAAB = CountTotalTickerAAB + 1
                    OpenCloseArrayAAB [i] = OpenCloseDifferenceAAB
                    
                    Dim RowVolumeAAB As Long
                    RowVolumeAAB = Cells(7, i).Value
                    VolumeArrayAAB [i] = RowVolumeAAB

                End If
                
                For Each element In OpenCloseArrayAAB
                    DailyAverageDifferenceAAB = (DailyAverageDifferenceAAB + CDbl(element))
                
                FinalDailyAverageAAB = (DailyAverageDifferenceAAB / CountTotalTickerAAB)
                
                
                'AAC
                '=================================================
                If (StrComp(ElementValue, "AAC", vbTextCompare) = 0) Then
                
                    'Grab Open
                    Dim OpenValue As Double
                    OpenValue = Cells(3, i).Value
                    
                    'Grab close
                    Dim CloseValue As Double
                    CloseValue = Cells(3, i).Value
                    
                    'Function calculate difference(opne, close) for current date
                    Dim OpenCloseDifferenceAAC As Double
                    OpenCloseDifferenceAAC = (OpenValue - CloseValue)
                    
                    CountTotalTickerAAC = CountTotalTickerAAC + 1
                    OpenCloseArrayAAC [i] = OpenCloseDifferenceAAC
                    
                    Dim RowVolumeAAC As Long
                    RowVolumeAAC = Cells(7, i).Value
                    VolumeArrayAAC [i] = RowVolumeAAC

                End If
                
                For Each element In OpenCloseArrayAAC
                    DailyAverageDifferenceAAC = (DailyAverageDifferenceAAC + CDbl(element))
                
                FinalDailyAverageAAC = (DailyAverageDifferenceAAC / CountTotalTickerAAC)
                
                
                'AAN
                '=================================================
                
                If (StrComp(ElementValue, "AAN", vbTextCompare) = 0) Then
                
                    'Grab Open
                    Dim OpenValue As Double
                    OpenValue = Cells(3, i).Value
                    
                    'Grab close
                    Dim CloseValue As Double
                    CloseValue = Cells(3, i).Value
                    
                    'Function calculate difference(opne, close) for current date
                    Dim OpenCloseDifference As Double
                    OpenCloseDifference = (OpenValue - CloseValue)
                    
                    CountTotalTickerAAN = CountTotalTickerAAN + 1
                    OpenCloseArrayAAN [i] = OpenCloseDifference
                    
                    Dim RowVolumeAAN As Long
                    RowVolumeAAN = Cells(7, i).Value
                    VolumeArrayAAN [i] = RowVolumeAAN

                End If
                
                For Each element In OpenCloseArrayAAN
                    DailyAverageDifferenceAAN = (DailyAverageDifferenceAAN + CDbl(element))
                
                FinalDailyAverageAAN = (DailyAverageDifferenceAAN / CountTotalTickerAAN)
                
                
                'AAP
                '=================================================
                If (StrComp(ElementValue, "AAP", vbTextCompare) = 0) Then
                
                    'Grab Open
                    Dim OpenValue As Double
                    OpenValue = Cells(3, i).Value
                    
                    'Grab close
                    Dim CloseValue As Double
                    CloseValue = Cells(3, i).Value
                    
                    'Function calculate difference(opne, close) for current date
                    Dim OpenCloseDifference As Double
                    OpenCloseDifference = (OpenValue - CloseValue)
                    
                    CountTotalTickerAAP = CountTotalTickerAAP + 1
                    OpenCloseArrayAAP [i] = OpenCloseDifference
                    
                    Dim RowVolumeAAP As Long
                    RowVolumeAAP = Cells(7, i).Value
                    VolumeArrayAAP [i] = RowVolumeAAP

                End If
                
                For Each element In OpenCloseArrayAAP
                    DailyAverageDifferenceAAP = (DailyAverageDifferenceAAP + CDbl(element))
                
                FinalDailyAverageAAP = (DailyAverageDifferenceAAP / CountTotalTickerAAP)
                
                
                'AAT
                '=================================================
                
                If (StrComp(ElementValue, "AAT", vbTextCompare) = 0) Then
                
                    'Grab Open
                    Dim OpenValue As Double
                    OpenValue = Cells(3, i).Value
                    
                    'Grab close
                    Dim CloseValue As Double
                    CloseValue = Cells(3, i).Value
                    
                    'Function calculate difference(opne, close) for current date
                    Dim OpenCloseDifference As Double
                    OpenCloseDifference = (OpenValue - CloseValue)
                    
                    CountTotalTickerAAT = CountTotalTickerAAT + 1
                    OpenCloseArrayAAT [i] = OpenCloseDifference
                    
                    Dim RowVolumeAAT As Long
                    RowVolumeAAT = Cells(7, i).Value
                    VolumeArrayAAT [i] = RowVolumeAAT

                End If
                
                For Each element In OpenCloseArrayAAT
                    DailyAverageDifferenceAAT = (DailyAverageDifferenceAAT + CDbl(element))
                
                FinalDailyAverageAAT = (DailyAverageDifferenceAAT / CountTotalTickerAAT)
                
                
                'AAV
                '=================================================
                If (StrComp(ElementValue, "AAV", vbTextCompare) = 0) Then
                
                    'Grab Open
                    Dim OpenValue As Double
                    OpenValue = Cells(3, i).Value
                    
                    'Grab close
                    Dim CloseValue As Double
                    CloseValue = Cells(3, i).Value
                    
                    'Function calculate difference(opne, close) for current date
                    Dim OpenCloseDifference As Double
                    OpenCloseDifference = (OpenValue - CloseValue)
                    
                    CountTotalTickerAAV = CountTotalTickerAAV + 1
                    OpenCloseArrayAAV [i] = OpenCloseDifference
                    
                    Dim RowVolumeAAV As Long
                    RowVolumeAAV = Cells(7, i).Value
                    VolumeArrayAAV [i] = RowVolumeAAV

                End If
                
                For Each element In OpenCloseArrayAAV
                    DailyAverageDifferenceAAV = (DailyAverageDifferenceAAV + CDbl(element))
                
                FinalDailyAverageAAV = (DailyAverageDifferenceAAV / CountTotalTickerAAV)
                
                
                'AB
                '=================================================
                If (StrComp(ElementValue, "AB", vbTextCompare) = 0) Then
                
                    'Grab Open
                    Dim OpenValue As Double
                    OpenValue = Cells(3, i).Value
                    
                    'Grab close
                    Dim CloseValue As Double
                    CloseValue = Cells(3, i).Value
                    
                    'Function calculate difference(opne, close) for current date
                    Dim OpenCloseDifference As Double
                    OpenCloseDifference = (OpenValue - CloseValue)
                    
                    CountTotalTickerAB = CountTotalTickerAB + 1
                    OpenCloseArrayAB [i] = OpenCloseDifference
                    
                    Dim RowVolumeAB As Long
                    RowVolumeAB = Cells(7, i).Value
                    VolumeArrayAB [i] = RowVolumeAB

                End If
                
                For Each element In OpenCloseArrayAB
                    DailyAverageDifferenceAB = (DailyAverageDifferenceAB + CDbl(element))
                
                FinalDailyAverageAB = (DailyAverageDifferenceAB / CountTotalTickerAB)
                
                
                'ABB
                '=================================================
                
                If (StrComp(ElementValue, "ABB", vbTextCompare) = 0) Then
                
                    'Grab Open
                    Dim OpenValue As Double
                    OpenValue = Cells(3, i).Value
                    
                    'Grab close
                    Dim CloseValue As Double
                    CloseValue = Cells(3, i).Value
                    
                    'Function calculate difference(opne, close) for current date
                    Dim OpenCloseDifference As Double
                    OpenCloseDifference = (OpenValue - CloseValue)
                    
                    CountTotalTickerABB = CountTotalTickerABB + 1
                    OpenCloseArrayABB [i] = OpenCloseDifference
                    
                    Dim RowVolumeABB As Long
                    RowVolumeABB = Cells(7, i).Value
                    VolumeArrayABB [i] = RowVolumeABB

                End If
                
                For Each element In OpenCloseArrayABB
                    DailyAverageDifferenceABB = (DailyAverageDifferenceABB + CDbl(element))
                
                FinalDailyAverageABB = (DailyAverageDifferenceABB / CountTotalTickerABB)
                
                
                'ABBV
                '=================================================
                If (StrComp(ElementValue, "ABBV", vbTextCompare) = 0) Then
                
                    'Grab Open
                    Dim OpenValue As Double
                    OpenValue = Cells(3, i).Value
                    
                    'Grab close
                    Dim CloseValue As Double
                    CloseValue = Cells(3, i).Value
                    
                    'Function calculate difference(opne, close) for current date
                    Dim OpenCloseDifference As Double
                    OpenCloseDifference = (OpenValue - CloseValue)
                    
                    CountTotalTickerABBV = CountTotalTickerABBV + 1
                    OpenCloseArrayABBV [i] = OpenCloseDifference
                    
                    Dim RowVolumeABBV As Long
                    RowVolumeABBV = Cells(7, i).Value
                    VolumeArrayABBV [i] = RowVolumeABBV

                End If
                
                For Each element In OpenCloseArrayABBV
                    DailyAverageDifferenceABBV = (DailyAverageDifferenceABBV + CDbl(element))
                
                FinalDailyAverageABBV = (DailyAverageDifferenceABBV / CountTotalTickerABBV)
                
                
                'ABC
                '=================================================
                If (StrComp(ElementValue, "ABC", vbTextCompare) = 0) Then
                
                    'Grab Open
                    Dim OpenValue As Double
                    OpenValue = Cells(3, i).Value
                    
                    'Grab close
                    Dim CloseValue As Double
                    CloseValue = Cells(3, i).Value
                    
                    'Function calculate difference(opne, close) for current date
                    Dim OpenCloseDifference As Double
                    OpenCloseDifference = (OpenValue - CloseValue)
                    
                    CountTotalTickerABC = CountTotalTickerABC + 1
                    OpenCloseArrayABC [i] = OpenCloseDifference
                    
                    Dim RowVolumeABC As Long
                    RowVolumeABC = Cells(7, i).Value
                    VolumeArrayABC [i] = RowVolumeABC

                End If
                
                For Each element In OpenCloseArrayABC
                    DailyAverageDifferenceABC = (DailyAverageDifferenceABC + CDbl(element))
                
                FinalDailyAverageABC = (DailyAverageDifferenceABC / CountTotalTickerABC)
                
                
            TickerCount = (TickerCount + 1)
            'Display ColumnCellTotal
        
        Next i
        
        
            
        '=================================================================================
        'OUTPUT
        
        'Create the Table that will house the collected volume
        '=================================================
          
        ' Declares the first Cell of the new Table
        Dim TableOrigin As String
        
        TableOrigin = ws.Cells(1, (Columns.Count + 2))
        
        TableOrigin.Value = "Stock Market Analysis"
        
        'Table Titles
        '==============================================
        
        Dim TickerTitle As String
        TickerTitle = "Ticker"
        ws.Cells(2, (Columns.Count + 2)) = TickerTitle
        
        Dim YearlyChange As String
        YearlyChange = "Yearly Change"
        ws.Cells(2, (Columns.Count + 3)) = YearlyChange
        
        Dim PercentChange As String
        PercentChange = "Percent Change"
        ws.Cells(2, (Columns.Count + 4)) = PercentChange
        
        Dim TotalStockVolume As String
        TotalStockVolume = "Total Stock Volume"
        ws.Cells(2, (Columns.Count + 5)) = TotalStockVolume
                
        Dim AverageDailyChange As String
        AverageDailyChange = "Average Daily Change Per Year"
        ws.Cells(2, (Columns.Count + 6)) = AverageDailyChange
        
        
        '=================================================
        'OUTPUT
        
        'With table outlined now fill in the data for each DataSet
        
        '=================================================
        'Ticker CHECK
        'YearlyChange CHECK
        'PercentChange
        'TotalStockValue CHECK
        'AverageDailyChangePerYear CHECK
        '=================================================
        
        Dim TotalCountTicker As Long
        
        'CountTotalTickerA
        Dim ATitle As String
        ATitle = "A"
        ws.Cells(3, (Columns.Count + 2)) = ATitle
        
        'YearlyChangeA
        Dim YearOpenA As Double
        Dim YearCloseA As Double
        YearOpenA = Cells(2, 3).Value
        TotalCountTicker = (CountTotalTickerA + 1)
        YearCloseA = Cells(TotalCountTicker, 6).Value
        
        YearlyChangeA = (YearOpenA - YearCloseA)
        ws.Cells(3, (Columns.Count + 3)) = YearlyChangeA
        
        'TotalStockVolumeA
        For Each element In VolumeArrayA
            TotalStockVolumeA = (TotalStockVolumeA + CLng(element))
        ws.Cells(3, (Columns.Count + 5)) = TotalStockVolumeA
        
        GreatestTotalVolume [0] = TotalStockVolumeA
        GreatestTotalVolumeE [0] = ATitle
        
        'AverageDailyChangePerYear
        ws.Cells(3, (Columns.Count + 6)) = FinalDailyAverageA
        '=================================================
        
        
        'CountTotalTickerAA
        Dim AATitle As String
        AATitle = "AA"
        ws.Cells(4, (Columns.Count + 2)) = AATitle
        
        'YearlyChangeAA
        Dim YearOpenAA As Double
        Dim YearCloseAA As Double
        YearOpenAA = Cells(2, 3).Value
        TotalCountTicker = (CountTotalTickerAA)
        YearCloseAA = Cells(TotalCountTicker, 6).Value
        
        YearlyChangeAA = (YearOpenAA - YearCloseAA)
        ws.Cells(4, (Columns.Count + 3)) = YearlyChangeAA
        
        'TotalStockVolumeAA
        For Each element In VolumeArrayAA
            TotalStockVolumeAA = (TotalStockVolumeAA + CLng(element))
        ws.Cells(4, (Columns.Count + 5)) = TotalStockVolumeAA
        
        GreatestTotalVolume [1] = TotalStockVolumeAA
        GreatestTotalVolumeE [1] = AATitle
        
        'AverageDailyChangePerYear
        ws.Cells(4, (Columns.Count + 6)) = FinalDailyAverageAA
        '=================================================
        
        
        'CountTotalTickerAAB
        Dim AABTitle As String
        AABTitle = "AAB"
        ws.Cells(5, (Columns.Count + 2)) = AABTitle
        
        'YearlyChangeAAB
        Dim YearOpenAAB As Double
        Dim YearCloseAAB As Double
        YearOpenAAB = Cells(2, 3).Value
        TotalCountTicker = (CountTotalTickerAAB)
        YearCloseAAB = Cells(TotalCountTicker, 6).Value
        
        YearlyChangeAAB = (YearOpenAAB - YearCloseAAB)
        ws.Cells(5, (Columns.Count + 3)) = YearlyChangeAAB
        
        'TotalStockVolumeAAB
        For Each element In VolumeArrayAAB
            TotalStockVolumeAAB = (TotalStockVolumeAAB + CLng(element))
        ws.Cells(5, (Columns.Count + 5)) = TotalStockVolumeAAB
        
        GreatestTotalVolume [2] = TotalStockVolumeAAB
        GreatestTotalVolumeE [2] = AABTitle
        
        'AverageDailyChangePerYear
        ws.Cells(5, (Columns.Count + 6)) = FinalDailyAverageAAB
        '=================================================
        
        
        'CountTotalTickerAAC
        Dim AACTitle As String
        AACTitle = "AAC"
        ws.Cells(6, (Columns.Count + 2)) = AACTitle
        
        'YearlyChangeAAC
        Dim YearOpenAAC As Double
        Dim YearCloseAAC As Double
        YearOpenAAC = Cells(2, 3).Value
        TotalCountTicker = (CountTotalTickerAAC)
        YearCloseAAC = Cells(TotalCountTicker, 6).Value
        
        YearlyChangeAAC = (YearOpenAAC - YearCloseAAC)
        ws.Cells(6, (Columns.Count + 3)) = YearlyChangeAAC
        
        'TotalStockVolumeAAC
        For Each element In VolumeArrayAAC
            TotalStockVolumeAAC = (TotalStockVolumeAAC + CLng(element))
        ws.Cells(6, (Columns.Count + 5)) = TotalStockVolumeAAC
        
        GreatestTotalVolume [3] = TotalStockVolumeAAC
        GreatestTotalVolumeE [3] = AACTitle
        
        'AverageDailyChangePerYear
        ws.Cells(6, (Columns.Count + 6)) = FinalDailyAverageAAC
        '=================================================
        
        
        'CountTotalTickerAAN
        Dim AANTitle As String
        AANTitle = "AAN"
        ws.Cells(7, (Columns.Count + 2)) = AANTitle
        
        'YearlyChangeAAN
        Dim YearOpenAAN As Double
        Dim YearCloseAAN As Double
        YearOpenAAN = Cells(2, 3).Value
        TotalCountTicker = (CountTotalTickerAAN)
        YearCloseAAN = Cells(TotalCountTicker, 6).Value
        
        YearlyChangeAAN = (YearOpenAAN - YearCloseAAN)
        ws.Cells(7, (Columns.Count + 3)) = YearlyChangeAAN
        
        'TotalStockVolumeAAN
        For Each element In VolumeArrayAAN
            TotalStockVolumeAAN = (TotalStockVolumeAAN + CLng(element))
        ws.Cells(7, (Columns.Count + 5)) = TotalStockVolumeAAN
        
        GreatestTotalVolume [4] = TotalStockVolumeAAN
        GreatestTotalVolumeE [4] = AANTitle
        
        'AverageDailyChangePerYear
        ws.Cells(7, (Columns.Count + 6)) = FinalDailyAverageAAN
        '=================================================
        
        
        'CountTotalTickerAAP
        Dim AAPTitle As String
        AAPTitle = "AAP"
        ws.Cells(8, (Columns.Count + 2)) = AAPTitle
        
        'YearlyChangeAAP
        Dim YearOpenAAP As Double
        Dim YearCloseAAP As Double
        YearOpenAAP = Cells(2, 3).Value
        TotalCountTicker = (CountTotalTickerAAP)
        YearCloseAAP = Cells(TotalCountTicker, 6).Value
        
        YearlyChangeAAP = (YearOpenAAP - YearCloseAAP)
        ws.Cells(8, (Columns.Count + 3)) = YearlyChangeAAP
        
        'TotalStockVolumeAAP
        For Each element In VolumeArrayAAP
            TotalStockVolumeAAP = (TotalStockVolumeAAP + CLng(element))
        ws.Cells(8, (Columns.Count + 5)) = TotalStockVolumeAAP
        
        GreatestTotalVolume [5] = TotalStockVolumeAAP
        GreatestTotalVolumeE [5] = AAPTitle
        
        'AverageDailyChangePerYear
        ws.Cells(8, (Columns.Count + 6)) = FinalDailyAverageAAP
        '=================================================
        
        
        'CountTotalTickerAAT
        Dim AATTitle As String
        AATTitle = "AAT"
        ws.Cells(9, (Columns.Count + 2)) = AATTitle
        
        'YearlyChangeAAT
        Dim YearOpenAAT As Double
        Dim YearCloseAAT As Double
        YearOpenAAT = Cells(2, 3).Value
        TotalCountTicker = (CountTotalTickerAAT)
        YearCloseAAT = Cells(TotalCountTicker, 6).Value
        
        YearlyChangeAAT = (YearOpenAAT - YearCloseAAT)
        ws.Cells(9, (Columns.Count + 3)) = YearlyChangeAAT
        
        'TotalStockVolumeAAT
        For Each element In VolumeArrayAAT
            TotalStockVolumeAAT = (TotalStockVolumeAAT + CLng(element))
        ws.Cells(9, (Columns.Count + 5)) = TotalStockVolumeAAT
        
        GreatestTotalVolume [6] = TotalStockVolumeAAT
        GreatestTotalVolumeE [6] = AATTitle
        
        'AverageDailyChangePerYear
        ws.Cells(9, (Columns.Count + 6)) = FinalDailyAverageAAT
        '=================================================
        
        
        'CountTotalTickerAAV
        Dim AAVTitle As String
        AAVTitle = "AAV"
        ws.Cells(10, (Columns.Count + 2)) = AAVTitle
        
        'YearlyChangeAAV
        Dim YearOpenAAV As Double
        Dim YearCloseAAV As Double
        YearOpenAAV = Cells(2, 3).Value
        TotalCountTicker = (CountTotalTickerAAV)
        YearCloseAAV = Cells(TotalCountTicker, 6).Value
        
        YearlyChangeAAV = (YearOpenAAV - YearCloseAAV)
        ws.Cells(10, (Columns.Count + 3)) = YearlyChangeAAV
        
        'TotalStockVolumeAAV
        For Each element In VolumeArrayAAV
            TotalStockVolumeAAV = (TotalStockVolumeAAV + CLng(element))
        ws.Cells(10, (Columns.Count + 5)) = TotalStockVolumeAAV
        
        GreatestTotalVolume [7] = TotalStockVolumeAAV
        GreatestTotalVolumeE [7] = AAVTitle
        
        'AverageDailyChangePerYear
        ws.Cells(10, (Columns.Count + 6)) = FinalDailyAverageAAV
        '=================================================
        
        
        'CountTotalTickerAB
        Dim ABTitle As String
        ABTitle = "AB"
        ws.Cells(11, (Columns.Count + 2)) = ABTitle
        
        'YearlyChangeAB
        Dim YearOpenAB As Double
        Dim YearCloseAB As Double
        YearOpenAB = Cells(2, 3).Value
        TotalCountTicker = (CountTotalTickerAB)
        YearCloseAB = Cells(TotalCountTicker, 6).Value
        
        YearlyChangeAB = (YearOpenAB - YearCloseAB)
        ws.Cells(11, (Columns.Count + 3)) = YearlyChangeAB
        
        'TotalStockVolumeAB
        For Each element In VolumeArrayAB
            TotalStockVolumeAB = (TotalStockVolumeAB + CLng(element))
        ws.Cells(11, (Columns.Count + 5)) = TotalStockVolumeAB
        
        GreatestTotalVolume [8] = TotalStockVolumeAB
        GreatestTotalVolumeE [8] = ABTitle
        
        'AverageDailyChangePerYear
        ws.Cells(11, (Columns.Count + 6)) = FinalDailyAverageAB
        '=================================================
        
        
        'CountTotalTickerABB
        Dim ABBTitle As String
        ABBTitle = "ABB"
        ws.Cells(12, (Columns.Count + 2)) = ABBTitle
        
        'YearlyChangeABB
        Dim YearOpenABB As Double
        Dim YearCloseABB As Double
        YearOpenABB = Cells(2, 3).Value
        TotalCountTicker = (CountTotalTickerABB)
        YearCloseABB = Cells(TotalCountTicker, 6).Value
        
        YearlyChangeABB = (YearOpenABB - YearCloseABB)
        ws.Cells(12, (Columns.Count + 3)) = YearlyChangeABB
        
        'TotalStockVolumeABB
        For Each element In VolumeArrayABB
            TotalStockVolumeABB = (TotalStockVolumeABB + CLng(element))
        ws.Cells(12, (Columns.Count + 5)) = TotalStockVolumeABB
        
        GreatestTotalVolume [9] = TotalStockVolumeABB
        GreatestTotalVolumeE [9] = ABBTitle
        
        'AverageDailyChangePerYear
        ws.Cells(12, (Columns.Count + 6)) = FinalDailyAverageABB
        '=================================================
        
        
        'CountTotalTickerABBV
        Dim ABBVTitle As String
        ABBVTitle = "ABBV"
        ws.Cells(13, (Columns.Count + 2)) = ABBVTitle
        
        'YearlyChangeABBV
        Dim YearOpenABBV As Double
        Dim YearCloseABBV As Double
        YearOpenABBV = Cells(2, 3).Value
        TotalCountTicker = (CountTotalTickerABBV)
        YearCloseABBV = Cells(TotalCountTicker, 6).Value
        
        YearlyChangeABBV = (YearOpenABBV - YearCloseABBV)
        ws.Cells(13, (Columns.Count + 3)) = YearlyChangeABBV
        
        'TotalStockVolumeABBV
        For Each element In VolumeArrayABBV
            TotalStockVolumeABBV = (TotalStockVolumeABBV + CLng(element))
        ws.Cells(13, (Columns.Count + 5)) = TotalStockVolumeABBV
        
        GreatestTotalVolume [10] = TotalStockVolumeABBV
        GreatestTotalVolumeE [10] = ABBVTitle
        
        'AverageDailyChangePerYear
        ws.Cells(13, (Columns.Count + 6)) = FinalDailyAverageABBV
        '=================================================
        
        
        'CountTotalTickerABC
        Dim ABCTitle As String
        ABCTitle = "ABBV"
        ws.Cells(14, (Columns.Count + 2)) = ABCTitle
        
        'YearlyChangeABC
        Dim YearOpenABC As Double
        Dim YearCloseABC As Double
        YearOpenABC = Cells(2, 3).Value
        TotalCountTicker = (CountTotalTickerABC)
        YearCloseABC = Cells(TotalCountTicker, 6).Value
        
        YearlyChangeABC = (YearOpenABC - YearCloseABC)
        ws.Cells(14, (Columns.Count + 3)) = YearlyChangeABC
        
        'TotalStockVolumeABC
        For Each element In VolumeArrayABC
            TotalStockVolumeABC = (TotalStockVolumeABC + CLng(element))
        ws.Cells(14, (Columns.Count + 5)) = TotalStockVolumeABC
        
        GreatestTotalVolume [11] = TotalStockVolumeABC
        GreatestTotalVolumeE [11] = ABCTitle
        
        'AverageDailyChangePerYear
        ws.Cells(14, (Columns.Count + 6)) = FinalDailyAverageABC
        '=================================================================================
        'OUTPUT
            
        'Create the Table that will house the added elements
        '=================================================
        
        ws.Cells(2, (Columns.Count + 10)) = TickerTitle
        
        Dim ValueTitel As String
        ValueTitel = "Value"
        ws.Cells(2, (Columns.Count + 11)) = Value
        
        Dim GreatestIncrease As String
        GreatestIncrease = "Greatest % Increase"
        ws.Cells(3, (Columns.Count + 9)) = GreatestIncrease
        
        Dim GreatestDecrease As String
        GreatestDecrease = "Greatest % Degrease"
        ws.Cells(4, (Columns.Count + 9)) = GreatestIncrease
        
        Dim GreatestTotalVolume As String
        GreatestTotalVolume = "Greatest Total Volume"
        ws.Cells(5, (Columns.Count + 9)) = GreatestTotalVolume
        
        '=================================================
        'OUTPUT
        
        'With table outlined now fill in the data for the second DataSet
        '=================================================
        
        'Greatest % Increase
        'Greatest % Decrease
        'Greatest Total Volume
        '=================================================
        
        Dim GreatestPercentIncreaseElement As String
        Dim GreatestPercentIncreaseValue As Double
        
        Dim GreatestPercentDecreaseElement As String
        Dim GreatestPercentDecreaseValue As Double
        
        
        'Greatest Total Volume
        '=================================================
        Dim GreatestTotalVolumeElement As String
        Dim GreatestTotalVolumeValue As Long
        Dim ElementCount As Integer
        
        ElementCount = 0
        GreatestTotalVolumeValue = 0
        For Each element In GreatestTotalVolume
            
            If GreatestTotalVolumeValue < CLng(element) Then
                GreatestTotalVolumeValue = CLng(element)
                'GreatestTotalVolumeElement = GreatestTotalVolumeE[ElementCount]
            End If
            ElementCount = ElementCount + 1
        ws.Cells(5, (Columns.Count + 11)) = GreatestTotalVolumeValue
        ws.Cells(5, (Columns.Count + 10)) = GreatestTotalVolumeElement
        
        '=================================================
    
End Sub
