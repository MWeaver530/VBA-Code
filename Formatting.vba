sub AddNewColumns()
  dim colOrder as Variant
  dim renameArr as Variant
  dim indx as long
  dim lcol as long
  dim colSpec as string
  
  'Define Column order here
  colOrder = array("col1", "col2", "col3")
  
  'Rename columns
  'col1 = Column 1, col2 = Column 2, col3 = Column 3
  lcol = ActiveSheet.Cells(1, Columns.count).End(xlToLeft).column
  for x = 1 to lcol
    Select Case Cells(1, x)
      Case "col1"
        Cells(1, x) = "Column 1"
      Case "col2"
        Cells(1, x) = "Column 2"
      Case "col3"
        Cells(1, x) = "Column 3"
    end select
  next

  'Set Specificed Column to sort from
  colSpec = "C"
  for indx = LBound(colOrder) to UBound(colOrder)
    set search = Rows("1:1").Find(colOrder(indx), LookIn:=xlValues, Lookat:=xlWhole, _
      SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
    if search = nothing then
      Columns(colSpec).Insert Shift:=xlToright, CopyOrigin:=xlformatFromLeftOrAbove
      Columns(colSpec).ColumnWidth = 20
      Range(colSpec & 1).value = colOrder(indx)
      AddFormatBorders colSpec
    end if
  next indx

end sub


sub AddformatBorders(col as string)
  with Columns(col)
    .Borders(xlInsideVertical).LineStyle = xlContinuous
    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    .BorderAround xlContinuous
  end with

  with Range(col & "1").Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColordarkl
    .TintAndShade = -0.149998474074526
    .PatternTintandShade = 0
  end with
end sub


sub OrderColumns()
  dim search as range
  dim cnt as Integer
  dim colOrder as Variant
  dim indx as Integer
  
  'Column Order reference
  'A = column1, B = column2, C = column3, D = column4, E = column5
  
  'Define Column Order
  colOrder = Array("Column 1", "Column 2", "Column 3", "Column 4", "Column 5")
  cnt = 1
  
  For indx = LBound(colOrder) to Ubound(colOrder)
    Set search = Rows("1:1").Find(colOrder(indx), LookIn:=xlValues, LookAt:=xlWhole, _
      SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
    If Not search Is Nothing Then
      If search.Column <> cnt Then
        search.EntireColumn.Cut
        Columns(cnt).Insert shift:=xlToRight
        Application.CutCopyMode = False
      end If
    end If
    cnt = cnt + 1
  next
end sub
