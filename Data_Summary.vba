Sub Data_Summary()
    Dim ws as worksheet



    for each ws in thisworkbook.worksheets
        Dim lastrow as long
        lastrow = ws.cells(ws.rows.count, "A").end(xlup).row


        'add headers
        ws.cells(1, 10).value = "Ticker"
        ws.cells(1, 11).value = "Quarterly Change"
        ws.cells(1, 12).value = "Percent Change"
        ws.cells(1, 13).value = "Total Stock Volume"


        'make headers bold
        ws.cells(1, 10).font.bold = true
        ws.cells(1, 11).font.bold = true
        ws.cells(1, 12).font.bold = true
        ws.cells(1, 13).font.bold = true


        'start in the second row of the analysis column
        Dim greatestincrease as double
        Dim greatestdecrease as double
        Dim greatestvolume as double


        greatestincrease = -inf
        greatestdecrease = inf
        greatestvolume = 0


        Dim tickerincrease as string
        Dim tickerdecrease as string
        Dim tickervolume as string


        Dim y as long
        y = 2 
        Dim x as long
        x = 2 


        'start in the second row of the original data
        do while y <= lastrow
            Dim ticker as string
            Dim openprice as double
            Dim closeprice as double
            Dim volume as double


            ticker = ws.cells(y, 1).value
            openprice = ws.cells(y, 3).value
            volume = 0


            do while y <= lastrow and ws.cells(y, 1).value = ticker
                closeprice = ws.cells(y, 6).value
                volume = volume + ws.cells(y, 7).value
                y = y + 1
            loop


            Dim quarterlychange as double
            Dim percentchange as double


            quarterlychange = closeprice - openprice
            if openprice <> 0 then
                percentchange = (quarterlychange / openprice) * 100
            else
                percentchange = 0
            end if


            'fill data in columns J to M
            ws.cells(x, 10).value = ticker
            ws.cells(x, 11).value = quarterlychange
            ws.cells(x, 12).value = percentchange / 100
            ws.cells(x, 13).value = volume


            'change cell color based on the value
            if quarterlychange > 0 then
                ws.cells(x, 11).interior.color = RGB(0, 176, 80)
                ws.cells(x, 11).font.color = RGB(255, 255, 255)
            elseif quarterlychange < 0 then
                ws.cells(x, 11).interior.color = RGB(192, 0, 0)
                ws.cells(x, 11).font.color = RGB(255, 255, 255)
            end if


            'format the "Quarterly Change" and "Percent Change"
            ws.cells(x, 11).numberformat = "0.00"
            ws.cells(x, 12).numberformat = "0.00%"


            'change the color of "Percent Change" if it's negative
            if percentchange < 0 then
                ws.cells(x, 12).font.color = RGB(192, 0, 0)
            end if


            'check "Greatest % Increase" and "Greatest % Decrease"
            if percentchange > greatestincrease then
                greatestincrease = percentchange
                tickerincrease = ticker
            end if


            if percentchange < greatestdecrease then
                greatestdecrease = percentchange
                tickerdecrease = ticker
            end if


            'check "Greatest Total Volume"
            if volume > greatestvolume then
                greatestvolume = volume
                tickervolume = ticker
            end if


            x = x + 1
        loop


        'adjust the column size
        ws.columns("J:M").autofit


        'add headers and results
        ws.cells(1, 17).value = "Ticker"
        ws.cells(1, 18).value = "Value"


        ws.cells(2, 16).value = "Greatest % increase"
        ws.cells(3, 16).value = "Greatest % decrease"
        ws.cells(4, 16).value = "Greatest total volume"


        ws.cells(2, 17).value = tickerincrease
        ws.cells(3, 17).value = tickerdecrease
        ws.cells(4, 17).value = tickervolume


        ws.cells(2, 18).value = greatestincrease / 100
        ws.cells(3, 18).value = greatestdecrease / 100
        ws.cells(4, 18).value = greatestvolume


        ws.cells(1, 16).font.bold = true
        ws.cells(1, 17).font.bold = true
        ws.cells(1, 18).font.bold = true


        'format the "Percent Change" with two decimals
        ws.cells(2, 18).numberformat = "0.00%"
        ws.cells(3, 18).numberformat = "0.00%"


        'adjust the column size
        ws.columns("P:R").autofit


    next ws



End Sub
