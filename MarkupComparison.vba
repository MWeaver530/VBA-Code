sub MarkupComparisonWorkbook(comparisonFolder as String, comparisonFile as String)
  dim lrow as long
  dim lTyp as string, rTyp as string
  dim lCtm as range, rCtm as range
  dim lAsm as range, rAsm as range
  dim lQty as range, rQty as range
  dim lPrt as range, rPrt as range

  Application.ScreenUpdating = False
  Application.DisplayAlerts = False

  Workbooks.Open fileName:=comparisonFolder & comparisonFile & ".xlsx"

  lrow = ActiveSheet.Cells(rows.Count, 1).End(xlUp).Row

  for x = 3 to lrow
    set lAsm = range("A" & x)
    set lPrt = range("B" & x)
    set lQty = range("C" & x)
    set lCtm = range("D" & x)
    set rCtm = range("E" & x)
    set rAsm = range("F" & x)
    set rPrt = range("G" & x)
    set rQty = range("H" & x)

    'Check if part is only on one BOM
    if Not Isempty(lAsm.value) And IsEmpty(rAsm.value) then
      lCtm.value = lCtm.value & "Not on Build"
      lCtm.Interior.Color = RGB(255, 255, 0)
    end if
    if IsEmpty(lAsm.value) And Not IsEmpty(rAsm.value) then
      rCtm.value = rCtm.vlue & "Not on Build"
      rCtm.Interior.Color = RGB(255, 255, 0)
    end if

    'check Quantities
    if lQty.value > rQty.value then
      lCtm.value = lCtm.value & "QTY!, "
      lCtm.Interior.Color = RGB(255, 255, 0)
      lQty.Interior.Color = RGB(255, 255, 0)
    end if
    if rQty.value > lQty.value then
      rCtm.value = rCtm.value & "QTY!, "
      rCtm.Interior.Color = RGB(255, 255, 0)
      rQty.Interior.Color = RGB(255, 255, 0)
    end if

    'Check Part
    if Not Isempty(lPrt) And Not IsEmpty(rPrt) then
      if lPrt.value <> rPrt.value then
        lPrt.Interior.Color = RGB(255, 255, 0)
        rPrt.Interior.Color = RGB(255, 255, 0)
      end if
    end if
  next

  ActiveWorkbook.Save
  ActiveWorkbook.Close
            
  Application.ScreenUpdating = true
  Application.DisplayAlerts = true
end sub
