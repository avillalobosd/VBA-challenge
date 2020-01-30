Sub stock()


'Valores del Loop
Dim openi, highi, lowi, closei As Double
Dim openf, highf, lowf, closef As Double
Dim sumVol As Variant
Dim ticker As String



For Each ws In Worksheets


            ticker = ws.Cells(2, 1).Value
            openi = ws.Cells(2, 3).Value
            highi = ws.Cells(2, 4).Value
            lowi = ws.Cells(2, 5).Value
            closei = ws.Cells(2, 6).Value
            sumVol = ws.Cells(2, 7).Value
            ' MsgBox ("open " + Str(openi) + " highi " + Str(highi) + " Lowi " + Str(lowi) + " closei " + Str(closei))



            'Cantidad de Filas
            Dim limitRow As Long
            limitRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

            'Posicion de Impresion de Datos
            Dim showTableR
            showTableR = 2

            'Encabezado de Tablas Impresión
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"



               ' Recorre Toda la hoja en posición Ticker
        
        
        
                    For i = 3 To limitRow + 1

                        If ticker <> ws.Cells(i, 1).Value Then
            
                            ' Imprime datos
                           ws.Cells(showTableR, 9).Value = ticker
                           ws.Cells(showTableR, 10).Value = closef - openi
                        If (openi = 0) Then
                           ws.Cells(showTableR, 11).Value = 0
                         Else
                 
                           ws.Cells(showTableR, 11).Value = (closef / openi) - 1
                         End If
                
                         ws.Cells(showTableR, 12).Value = sumVol
                
                         If (ws.Cells(showTableR, 10).Value < 0) Then
                         ws.Cells(showTableR, 10).Interior.Color = RGB(255, 0, 0)
                         Else
                         ws.Cells(showTableR, 10).Interior.Color = RGB(0, 255, 0)
                         End If
                
                
                        'Set de nuevos datos iniciales
                        ticker = ws.Cells(i, 1).Value
                        openi = ws.Cells(i, 3).Value
                        highi = ws.Cells(i, 4).Value
                        lowi = ws.Cells(i, 5).Value
                        closei = ws.Cells(i, 6).Value
                        sumVol = ws.Cells(i, 7).Value
                        
                        
                        showTableR = showTableR + 1
                    
                        ElseIf ticker = ws.Cells(i, 1).Value Then
                        
                        
                            openf = ws.Cells(i, 3).Value
                            highf = ws.Cells(i, 4).Value
                            lowf = ws.Cells(i, 5).Value
                            closef = ws.Cells(i, 6).Value
                            sumVol = sumVol + ws.Cells(i, 7).Value
                        
                        
                        End If
            
                    Next i
            
            
Next ws
            
For Each ws In Worksheets
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        limitRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        MsgBox (limitRow)
        
        Dim TickerI, TickerD, TickerV As String
        Dim vI, vD, vV As Variant
        
        TickerI = ws.Cells(2, 9)
        TickerD = ws.Cells(2, 9)
        TickerV = ws.Cells(2, 9)
        
        vI = ws.Cells(2, 11)
        vD = ws.Cells(2, 11)
        vV = ws.Cells(2, 12)
        
        For i = 3 To limitRow + 1
        
        If ws.Cells(i, 11).Value > vI Then
        
        vI = ws.Cells(i, 11).Value
        TickerI = ws.Cells(i, 9).Value
        
        End If
        
        If ws.Cells(i, 11).Value < vD Then
        
        vD = ws.Cells(i, 11).Value
        TickerD = ws.Cells(i, 9).Value
        
        End If
        
        If ws.Cells(i, 12).Value > vV Then
        
        vV = ws.Cells(i, 12).Value
        TickerV = ws.Cells(i, 9).Value
        
        End If
        
        
        
        Next i
        
        ws.Cells(2, 16).Value = TickerI
        ws.Cells(2, 17).Value = vI
        
        ws.Cells(3, 16).Value = TickerD
        ws.Cells(3, 17).Value = vD
        
        ws.Cells(4, 16).Value = TickerV
        ws.Cells(4, 17).Value = vV



Next ws

End Sub
