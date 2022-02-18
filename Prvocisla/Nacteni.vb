Imports xls = Microsoft.Office.Interop.Excel
Module Nacteni

    Sub NacteniSoub(soubor As String)
        Dim excelApp As New xls.Application
        Dim sesit_1 As xls.Workbook
        Dim list_1 As xls.Worksheet
        Dim vychoziBunka, aktualniBunka As xls.Range
        Dim oblast_bunek_1 As xls.Range
        Dim pocetRadku, jeInteger, cislo, k As Integer
        Dim text As String = ""
        Dim prvocislo As Boolean
        k = 0

        sesit_1 = excelApp.Workbooks.Open(soubor)

        list_1 = sesit_1.Worksheets(1)
        Do
            k += 1
            vychoziBunka = list_1.Cells(1, k)
        Loop While vychoziBunka.Value = ""
        oblast_bunek_1 = vychoziBunka.CurrentRegion
        pocetRadku = oblast_bunek_1.Rows.Count
        For i = 1 To pocetRadku
            prvocislo = False
            aktualniBunka = oblast_bunek_1(i, 1)
            If Integer.TryParse(aktualniBunka.Value, jeInteger) Then
                cislo = Integer.Parse(aktualniBunka.Value)
                If cislo > 0 Then
                    prvocislo = Primenumbers(cislo)
                    If prvocislo = True Then
                        text += aktualniBunka.Value + Chr(10)
                    End If
                End If
            End If
        Next

        sesit_1.Close()
        excelApp.Quit()
        excelApp = Nothing

        MsgBox(text,, "Prvočísla v souboru jsou:")

    End Sub
End Module
