sub wallstreethard():
'create worksheet variable
    dim ws As worksheet

    'create a for each loop
    for each ws in worksheets

        'create column headers for new variables
        ws.Range("I1") = "Ticker Name"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Volume"

        'format greatest table
        ws.Range("O2").value = "Greatest % Increase"
        ws.Range("O3").value = "Greatest % Decrease"
        ws.Range("O4").value = "Greatest Total Volume"
        ws.Range("P1").value = "Ticker"
        ws.Range("Q1").value = "Value"

        'set an initial variable for holding the ticker name
        dim tickername as string

        'set an initial variable for holding the total volume per ticker
        dim volumetotal as double
        volumetotal = 0

        'set an initial variable for holding yearly change per ticker
        dim yearchange as double

        'keep track of the location for each ticker in the summary table
        dim summarytablerow as integer
        summarytablerow = 2

        'set initial variable to hold open value for ticker
        dim openvalue as double
    
        'set initial variable to hold close value for ticker
        dim closevalue as double

        'set percent change variable
        dim percentchange as double

        'set greatest percent increase variable
        dim greatestpercentincrease as double
        greatestpercentincrease = 0

        'set greatest percent decrease variable
        dim greatestpercentdecrease as double
        greatestpercentdecrease = 0

        'set greatest volume change variable
        dim greatesttotalvolume as long
        greatesttotalvolume = 0

        'create last row variable 
        dim lastrow as long
        lastrow = ws.cells(rows.count, 1).end(xlup).row

            'loop through rows
            For i = 2 to lastrow

                'check if we are still within the same ticker
                 if ws.cells(i + 1, 1).Value <> ws.cells(i, 1).Value then 

                'set the ticker name
                tickername = ws.cells(i, 1).Value

                'create ticker count
                tickercount = Application.WorksheetFunction.CountIf(Range("A:A"), Cells(i, 1).Value)
                
                'add to the volume total
                volumetotal = volumetotal + ws.cells(i, 7).Value

                    'find company with greatest volume total
                    if volumetotal > greatesttotalvolume then
                    greatesttotalvolume = volumetotal

                    'print value of greatest total volume to greatest table
                     ws.Range("Q4").value = greatesttotalvolume

                    'print corresponding ticker name
                    ws.Range("P4").value = ws.cells(i, 1).value
                
                     end if

                'find open value
                openvalue = ws.cells(i - tickercount + 1, 3)
            
                'find closing value
                closevalue = ws.cells(i, 6).Value

                'calculate yearly change
                yearchange = closevalue - openvalue

                'calculate percent change 
                percentchange = (1-(closevalue/openvalue)) * 100

                    'find company with greatest % increase
                    if percentchange > greatestpercentincrease then
                    greatestpercentincrease = percentchange

                    'find compant with greatest % decrease
                    elseif percentchange < greatestpercentdecrease then
                    greatestpercentdecrease = percentchange

                    end if

                'print value of greatest % increase to greatest table
                ws.Range("Q2").value = greatestpercentincrease & "%"
            
                'print correlating ticker to greatest table
                ws.Range("P2").value = cells(i, 1).value

                'print value of greatest % decrease to greatest table
                ws.Range("Q3").value = greatestpercentdecrease & "%"

                'print correlating ticker to greatest table
                ws.Range("P3").value = cells(i,1).value

                'print ticker name in the summary table
                ws.Range("I" & summarytablerow).value = tickername

                'print yearly change in the summary table
                ws.Range("J" & summarytablerow).value = yearchange

                    'set negative changes to red
                    if yearchange < 0 then
                        ws.Range("J" & summarytablerow).interior.colorindex = 3
                
                    'set positive changes to green
                    else
                        ws.Range("J" & summarytablerow).interior.colorindex = 4
                    
                    end if

                'print percent change to the summary table
                ws.Range("K" & summarytablerow).value = percentchange & "%"

                'print the ticker volume to the summary table
                ws.Range("L" & summarytablerow).value = volumetotal

                'add one row to the summary table
                summarytablerow = summarytablerow + 1

                'reset the volume total
                volumetotal = 0

                'if the cell immediatley following a row is the same ticker
            else

                'add the volume total
                volumetotal = volumetotal + ws.cells(i, 3).value

            end if

        next i

    next ws

End sub