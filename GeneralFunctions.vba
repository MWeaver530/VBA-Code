function PullFileList(folderPath as string)
  dim vaArray as variant
  dim i as integer
  dim oFSO as object
  dim oFolder as object
  dim oFile as object
  dim oFiles as object

  set oFSO = CreateObject("Scripting.FileSystemObject")
  set oFolder = oFSO.GetFolder(folderpath)
  set oFiles = ofolder.Files

  if oFiles.count = 0 then Exit Function

  ReDim vaArray (1 to oFiles.count)
  i = 1
  for each oFile in ofiles
    vaArray(i) = oFile.name
    i = i + 1
  next

  set oFSO = nothing
  set oFolder = nothing
  set oFile = nothing
  set oFiles= nothing

  PullFileList = vaArray

end function
