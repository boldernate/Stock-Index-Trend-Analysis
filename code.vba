'****************************
Sub import_Nasdaq_100_data_set()
'
' copies the data from the active worksheet to a new sheet in the template
'
Set aw = ActiveSheet

Set tw = ThisWorkbook.Sheets.Add
tw.Name = "Nasdaq Data"

aw.UsedRange.Copy
tw.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats

End Sub
'****************************
Sub vlookup_Nasdaq_Tech_100_data()
'
' imports columns using vlookup if the values in column A are an exact match
'
Set tw = ThisWorkbook.Sheets("Nasdaq Data")
Set A = Application.GetOpenFilename()
Set aw = A.Sheets(1)

' Adds column headings to Nasdaq Data
tw.Range("G1") = "Tech Index Value"
tw.Range("H1") = "Tech High"
tw.Range("I1") = "Tech Low"
tw.Range("J1") = "Tech Total MV"
tw.Range("K1") = "Tech Dividend MV"

twlr = tw.Cells(Rows.Count, "A").End(xlUp).Row
awlr = aw.Cells(Rows.Count, "A").End(xlUp).Row

For I = 2 To twlr
    lv = tw.Cells(I, "A").Value
    Set rV = aw.Range("A2:F" & awlr)
    
    gresult = Application.VLookup(lv, rV, 2, False)
    hresult = Application.VLookup(lv, rV, 3, False)
    iresult = Application.VLookup(lv, rV, 4, False)
    jresult = Application.VLookup(lv, rV, 5, False)
    kresult = Application.VLookup(lv, rV, 6, False)

    tw.Cells(I, "G") = gresult
    tw.Cells(I, "H") = hresult
    tw.Cells(I, "I") = iresult
    tw.Cells(I, "J") = jresult
    tw.Cells(I, "K") = kresult
Next I

End Sub
Sub SP500_Year_End_Comparison()
Set A = ActiveWorkbook
Set aw = ActiveSheet

Set t = ThisWorkbook
Set tw = t.Sheets("S&P 500")

tw.Range("A1:G100").ClearContents
aw.UsedRange.Copy tw.Range("A1")

twlc = tw.Cells(1, Columns.Count).End(xlToLeft).Column
twlr = tw.Cells(Rows.Count, "A").End(xlUp).Row

tw.Cells(1, 2) = A.Name & " Data"
tw.Cells(twlc + 1) = tw.Cells(1, 2).Value & " % Change"

For I = 2 To twlr
  percentchange = (tw.Cells(I, twlc).Value - tw.Cells(I - 1, twlc).Value) / tw.Cells(I - 1, twlc).Value
  tw.Cells(I, twlc + 1) = percentchange
Next I

twlc = tw.Cells(1, Columns.Count).End(xlToLeft).Column
p = GetOpenFilename()
Set n = Workbooks.Open(p)
Set nw = n.Sheets(1)

tw.Cells(1, twlc + 1) = n.Name & " Data"

nwlr = nw.Cells(Rows.Count, "A").End(xlUp).Row

For I = 2 To twlr
    lv = tw.Cells(I, "A").Value
    Set Rng = nw.Range("A2:B" & nwlr) ' assumes each data set has lookup column as column a and values in column b
    ci = 2

    vresult = Application.VLookup(lv, Rng, ci, False)
    tw.Cells(I, twlc + 1) = vresult
Next I

'*********************************************************
' Adds percent change columns
twlc = tw.Cells(1, Columns.Count).End(xlToLeft).Column

For I = 2 To twlr
  percentchange = (tw.Cells(I, twlc).Value - tw.Cells(I - 1, twlc).Value) / tw.Cells(I - 1, twlc).Value
  tw.Cells(I, twlc + 1) = percentchange
Next I



End Sub
Sub twenty_four_Month_Trend()
Set tw = ThisWorkbook.Sheets.Add
tw.Name = "24 Month Trend"






End Sub
Sub Append_Daily_Data_Nasdaq_Sheet()
' Appends daily data columns to the nasdaq data set
Set aw = ActiveSheet
Set tw = ThisWorkbook.Sheets("Nasdaq Data")

twlc = tw.Cells(1, Columns.Count).End(xlToLeft).Column
twlr = tw.Cells(Rows.Count, "A").End(xlUp).Row

awlc = aw.Cells(1, Columns.Count).End(xlToLeft).Column
awlr = aw.Cells(Rows.Count, "A").End(xlUp).Row

tw.Cells(1, twlc + 1) = aw.Name



Set rV = aw.Range("A2:B" & awlr)


For I = 2 To twlr
  lv = "*" & tw.Cells(I, "A").Value & "*"
  Set rV = aw.Range("A2:B" & awlr)
  lkup = Application.VLookup(lv, rV, 2, False)
  tw.Cells(I, twlc + 1) = lkup
Next I







End Sub

Sub render_dates_24_months()
' renders eomonth formula to find the last day of each of the last 24 months
Set tw = ActiveSheet

j = 25

For I = 2 To 26

    tw.Cells(I, "A") = Application.EoMonth(Now, "-" & j)
    j = j - 1
Next I

End Sub

Sub Add_Data_Column_Monthly_Data()
Set aw = ActiveSheet
Set tw = ThisWorkbook.Sheets("Monthly Data")

twlc = tw.Cells(1, Columns.Count).End(xlToLeft).Column
twlr = tw.Cells(Rows.Count, "A").End(xlUp).Row

awlc = aw.Cells(1, Columns.Count).End(xlToLeft).Column
awlr = aw.Cells(Rows.Count, "A").End(xlUp).Row

For I = 2 To twlr
    Set r = aw.Columns(1).Find(tw.Cells(I, "B").Value)
    If Not r Is Nothing Then
      tw.Cells(I, twlc + 1) = aw.Cells(rn, awlc).Value
    Else
      tw.Cells(I, twlc + 1) = "Not Found"
    End If
Next I




End Sub

