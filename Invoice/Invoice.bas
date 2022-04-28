Option Explicit

Private Sub Workbook_Open()
With Faktura.ComboBox1
    .AddItem "Przelew"
    .AddItem "Gotówka"
End With

Dim j As Integer
Dim rng As range
Dim count_customers As Integer
For Each rng In Klienci.UsedRange
If Application.WorksheetFunction.IsText(rng) Then
count_customers = count_customers + 1
End If
Next rng

For j = 1 To count_customers
    Faktura.ComboBox2.AddItem Worksheets("Klienci").Cells(j, 1)
Next j


Dim rnge As range
Dim count_cars As Integer
For Each rnge In Samochody.UsedRange
If Application.WorksheetFunction.IsText(rnge) Then
count_cars = count_cars + 1
End If
Next rnge

For j = 1 To count_cars
    Faktura.ComboBox3.AddItem Worksheets("Samochody").Cells(j, 1)
Next j

With Faktura.ComboBox4
    .AddItem "Krajowy"
    .AddItem "Międzynarodowy"
End With

End Sub
