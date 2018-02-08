sub wallstreeteasy():

'create worksheet variable
dim ws as worksheet

    'create a for each loop
    for each ws in worksheets

        'create column headers for new variables
        ws.Range("I1") = "Ticker Name"
        ws.Range("J1") = "Total Volume"

        'set an initial variable for holding the ticker name
        dim tickername as string

        'set an initial variable for holding the total volume per ticker
        dim volumetotal as double
        volumetotal = 0

        'keep track of the location for each ticker in the summary table
        dim summarytablerow as integer
        summarytablerow = 2

        'create last row variable
        Dim lastrow As Long
        lastrow = ws.cells(Rows.Count, 1).End(xlUp).Row

            'loop through all credit card purchases
            For i = 2 to lastrow

              'check if we are still within the same ticker
               if ws.cells(i + 1, 1).Value <> ws.cells(i, 1).Value then 

                  'set the ticker name
                   tickername = ws.cells(i, 1).Value

                  'add to the volume total
                  volumetotal = volumetotal + ws.cells(i, 7).Value
        
                  'print ticker name in the summary table
                  ws.Range("I" & summarytablerow).value = tickername

                  'print the ticker volume to the summary table
                  ws.Range("J" & summarytablerow).value = volumetotal

                  'add one row to the summary table
                  summarytablerow = summarytablerow + 1

                  'rest the volume total
                  volumetotal = 0

            'if the cell immediatley following a row is the same ticker
            else

                'add the volume total
                volumetotal = volumetotal + cells(i, 3).value

            end if

        next i
    
    next ws

End sub
