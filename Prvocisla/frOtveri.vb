Imports System.Windows.Forms
Public Class frOtveri
    Dim soubor As DialogResult
    Private Sub frOtveri_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim otevreniOK As Boolean = False

        With Me.ofdOtevri
            .Filter = "Soubor Excel .xlsx|*.xlsx|Vše|*.*"
            .FilterIndex = 1
            .Title = "Výběr "".xlsx"" souboru"
            .CheckFileExists = True
        End With
        Do
            soubor = Me.ofdOtevri.ShowDialog
            If soubor = DialogResult.OK And Strings.Right(ofdOtevri.FileName, 4) = "xlsx" Then
                NacteniSoub(ofdOtevri.FileName)
                otevreniOK = True
            Else
                MsgBox("Vyberte prosím soubor ve formátu Excel 2007-2013 (XLSX).",, "Otevření se nezdařilo")
            End If
        Loop While otevreniOK = False
        Me.Close()
    End Sub
End Class