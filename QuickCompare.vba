sub QuickCompare
  dim col1 as Collection
  dim col2 as Collection
  dim lrow as Long
  dim e as Long, g as Long

  lrow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
  Set col1 = ReadCollection(lrow, "A", "B")

  lrow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
  Set col2 = ReadCollection(lrow, "C", "D")

  e = 2
  g = 2

  indx1 = 1
  for each ele in col1
    temp1 = split(ele1, "<>")

    indx2 = 1
    for each ele2 in col2
      temp2 = split(ele2, "<>"
        
      if temp1(0) = temp2(0) then
        if temp1(1) = temp2(1) then
          Range("E" & e).value  = temp1(0)
          Range("F" & e).value  = temp1(1)
          e = e + 1
        else
          Range("G" & g).value  = temp1(0)
          Range("H" & g).value  = temp1(1) & "<>" & temp2(1)
          e = e + 1
        end if
        
        col1.remove (indx1)
        col2.remove (indx2)

        indx1 = indx1 - 1
        indx2 = indx2 - 1

        exit for
      end if

      indx2 = indx2 + 1
    next
    indx1 = indx1 + 1
  next
  
  i = 2
  for each ele in col1
    temp = split(ele, "<>")
    range("I" & i).value = temp(0)
    range("J" & i).value = temp(1)
    i = i + 1
  next

  k = 2
  for each ele2 in col2
    temp = split(ele2, "<>")
    range("K" & k).value = temp(0)
    range("L" & k).value = temp(1)
    i = i + 1
  next

end sub

function ReadCollection (lrow as long, 1 as string, c2 as string) as collection
  dim col as Collection
  range(c1 & "2:" & c2 & lrow).sort key1:=c1 & "2:" & c1 & lrow, order1=xlAscending, Header:=xlNo

  set col = new collection
  for x = 2 to lrow
    t1 = range(c1 & x).value
    t2 = range(c2 & x).value
    value = t1 & "<>" & t2
    col.add value
  next
  set ReadCollection = col
end function

sub ClearQuickCompareSheet()
  dim lrow as long
  dim lrow2 as long
  
  lrow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
  lrow2 = ActiveSheet.Cells(Rows.Count, 3).End(xlUp).Row

  if lrow2 > lrow then lrow = lrow2

  for x = 2 to lrow
    range("A" & x).EntireRow.ClearContents
  next
end sub
