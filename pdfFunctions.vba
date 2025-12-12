Sub ModifyPDF ()
  Dim drawingsFolder As String
  Dim asmFolder As String, pdfAsmFolder As String
  Dim exportFile As String
  Dim pdfArr As Variant, asmArr As Variant, arr As Variant
    
  'Set Folder locations
  drawingsFolder = Sheets ("Worksheet"). Range ("C1"). value & "\01 PDE\"
  asmFolder = Sheets ("Worksheet"). Range ("C1"). value & "\04 Assemblies\"
  pdfAsmFolder = Sheets ("Worksheet"). Range ("C1"). value & "\05 PDF Assemblies\"
    
  'Pul1 Drawings Array, 01 PDE
  pdfArr = PullFileList (drawingsFolder)
  'Pull Assembly Array, 04 Asm
  asmArr = PullFileList (asmFolder)
    
  For Each dwg In pdfArr
    For Each asm In asmArr
      If Left (asm, 14) = Left (dwg, 14) Then
        Application.ScreenUpdating = False
        'Turn asm file into PDF
        AssemblyloPDF asmFolder, Left (asm, 14), pdfAsmFolder
      
        'Append asm pdf to dwg
        AppendDrawing drawingsFolder, Left (dwg, 17), pdfAsmFolder, Left (am
        Application. ScreenUpdating = True
      end if
    next
  next
end sub

Sub AppendDrawing (drawingsFolder As String, dwg As String, pdfAsmFolder As String)
  'Insert other pdf into primary file
  Dim arrayFilePaths () As Variant
  Set app = CreateObject ("Acroexch. app")
  
  arrayFilePaths = Array((drawingsFolder & dwg & ".pdf"), (pdfAsmFolder & asm &".pdf"))
  
  Set primaryDoc - CreateObject ("AcroExch. PDDoc")
  OK = primaryDoc.Open (arrayFilePaths (0))
  
  For arrayIndex = 1 To UBound (arrayFilePaths)
    numPages = primaryDoc.GetNumPages () - 1
    
    Set sourceDoc = CreateObject ("AcroExch. PDDoc")
    OK = sourceDoc.Open (arrayFilePaths (arrayIndex))
    
    numberOfPagesToInsert = sourceDoc.GetNumPages
    OK - primaryDoc. InsertPages (numPages, sourceDoc, 0, numberOfPagesToInsert)
    OK - primaryDoc.Save (PDSaveFull, arrayFilePaths (0))
    
    Set sourceDoc = Nothing
  Next arrayIndex
        
  Set primaryDoc = Nothing
  app.Exit
  Set app = Nothing
            
end sub
