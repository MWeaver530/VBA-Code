function PullFileDictionary(location as string, strTarget as string) as object
  dim coll as collection
  dim dict as object
  set dict = CreateObject("Scripting.Dictionary")
  dim dictKey as string
  dim target as string
  dim rev as string

  dim oFSO as object
  dim oFolder as object
  dim oFile as object
  dim oFiles as object
  set oFSO = CreateObject("Scripting.FilesystemObject")
  set oFolder = oFSO.GetFolder(location)
  set oFile = oFolder.Files

  if oFiles.count = 0 then exit function

  for each oFile in oFiles
    temp = split(oFile.name, ".")
    
    if temp(0) = strTarget then
      dictKey = temp(0) & "." & temp(1)
      target = temp(2)

      if temp(3) <> "pdf" then
        target = temp(2) & "." & temp(3)
      end if

      if dict.exists(dictKey) then
        set tempColl = dict(dictKey)
        tempColl.add target
        set coll = tempColl
        set dict(dictKey) = coll
      else
        set coll = new collection
        coll.add target
        dict.add dictKey, coll
      end if
                
    end if
  next

  set PullFileDictionary = dict

end function
              
