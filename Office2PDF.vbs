call main()

sub main()
    dim argv
    set argv = wscript.arguments
    if lcase(right(wscript.fullname,11)) = "wscript.exe" then
        if argv.count < 1 then
            msgbox "No Files Detected"
        else
            call distributor(fileList(argv))
        end if
    else
        call executor(argv(0), argv(1), argv(2))
    end if
    wscript.quit
end sub

sub distributor(byRef fList)
    On Error Resume Next
    dim runPPTApp, runWordApp, fType
    dim objPPT, objWord
    dim lockPres, lockDoc

    dim listPPT, listWord, fileName
    set listPPT = createObject("system.collections.arrayList")
    set listWord = createObject("system.collections.arrayList")
    for each fileName in fList
        select case fileType(fileName)
            case 1 listPPT.add fileName
            case 2 listWord.add fileName
        end select
    next

    runPPTApp = 0
    runWordApp = 0

    for each fileName in listPPT
        if runPPTApp = 0 then
            set objPPT = getObject(, "powerPoint.application")
            if err.number <> 0 then
                err.clear
            else
                ' TODO: other choices
                msgbox "PowerPoint will be closed"
                objPPT.quit
            end if
            set objPPT = createObject("powerPoint.application")
            runPPTApp = 1
            set lockPres = objPPT.presentations.add(false)
            ' process count
            lockPres.slides.add 1, 1
            lockPres.slides(1).name = cstr(listPPT.count)
            ' process synchronization (spin lock)
            lockPres.slides.add 2, 1
            call newProcess(fileName, 1, chr(34) & lockPres.name & chr(34))
        else
            call newProcess(fileName, 1, chr(34) & lockPres.name & chr(34))
        end if
    next

    for each fileName in listWord
        if runWordApp = 0 then
            set objWord = getObject(, "word.application")
            if err.number <> 0 then
                err.clear
            else
                ' TODO: other choices
                msgbox "Word will be closed"
                objWord.quit
            end if
            set objWord = createObject("word.application")
            runWordApp = 1
            set lockDoc = objWord.documents.add(, , , false)
            lockDoc.paragraphs(1).id = cstr(listWord.count)
            call newProcess(fileName, 2, chr(34) & lockDoc.name & chr(34))
        else
            call newProcess(fileName, 2, chr(34) & lockDoc.name & chr(34))
        end if
    next

    if runPPTApp = 1 then
        do
            ' double check
            if objPPT.presentations.count <= 1 and cint(lockPres.slides(1).name) = 0 then
                wscript.sleep(500)
                exit do
            end if
        loop
        lockPres.close
        set lockPres = nothing
        objPPT.quit
        set objPPT = nothing
    end if

    if runWordApp = 1 then
        do
            ' double check
            if objWord.documents.count <= 1 and listWord.count = lockDoc.comments.count then
                wscript.sleep(500)
                exit do
            end if
        loop
        lockDoc.close(0)
        set lockDoc = nothing
        objWord.quit
        set objWord = nothing
    end if
end sub

sub executor(byVal fileName, byVal fType, byVal pPara)
    wscript.echo fileName
    select case convertSingle(fileName, fType, pPara)
        case 0  wscript.echo "successfully converted"
        ' TODO: solve this error?
        case 1  wscript.echo "synchronization error"
        case -1 wscript.echo "unsupported file type"
    end select
end sub

sub newProcess(byVal fileName, byVal fType, byVal pPara)
    dim wshell
    set wshell = createObject("wscript.shell")
    wshell.run "cmd.exe /c cscript.exe //nologo " &_
    chr(34) & wscript.scriptfullname & chr(34) & chr(32) &_
    chr(34) & fileName & chr(34) & chr(32) &_
    fType & chr(32) & pPara
    set wshell = nothing
end sub

function convertSingle(byVal fileName, byVal fType, byVal pPara)
    select case fType
        case 1 convertSingle = pptConvert(fileName, pPara)
        case 2 convertSingle = docConvert(fileName, pPara)
        case else convertSingle = -1
    end select
end function

function pptConvert(byVal fileName, byVal pPara)
    On Error Resume Next
    dim objPPT
    dim lockPres
    do
        set objPPT = getObject(, "powerPoint.application")
        if err.number = 0 then
            wscript.echo "PowerPoint is running"
            err.clear
            exit do
        else
            err.clear
        end if
    loop
    wscript.echo "Converting"
    ' By index or by name?
    set lockPres = objPPT.presentations(pPara)
    if lockPres.readOnly = true then
        pptConvert = 1
        exit function
    end if
    set objPres = objPPT.presentations.open(fileName, true, , false)
    ' spin lock
    do
        lockPres.slides(2).delete
        if err.number = 0 then
            exit do
        else
            err.clear
        end if
    loop
    ' enter critical section
    objPres.saveAs cutFileName(fileName) & "pdf", 32
    objPres.close
    set objPres = nothing
    ' decrease process count
    lockPres.slides(1).name = cstr(cint(lockPres.slides(1).name) - 1)
    ' leave critical section
    lockPres.slides.add 2, 1
    wscript.echo "Presentation is closed"
    pptConvert = 0
end function

function docConvert(byVal fileName, byVal pPara)
    On Error Resume Next
    dim objWord
    dim lockDoc, lockPara
    do
        set objWord = getObject(,"word.application")
        if err.number = 0 then
            wscript.echo "Word is running"
            err.clear
            exit do
        else
            err.clear
        end if
    loop
    wscript.echo "Converting"
    set objDoc = objWord.documents.open(fileName, , true)
    objDoc.saveAs cutFileName(fileName) & "pdf", 17
    objDoc.close(0)
    set objDoc = Nothing
    wscript.echo "Document is closed"
    set lockDoc = objWord.documents(pPara)
    if lockDoc.readOnly = true then
        docConvert = 1
        exit function
    end if
    ' paragraphs.count scheme is replaced by comments.count
    ' lockDoc.paragraphs.add
    lockDoc.comments.add lockDoc.paragraphs(1).range, ""
    docConvert = 0
end function

function fileList(byRef argv)
    dim fList, fso
    set fList = createObject("system.collections.arrayList")
    set fso = createObject("scripting.fileSystemObject")

    for each fileName in argv
        if fso.folderExists(fileName) then
            call folderExpand(fileName, fList)
        else
            fList.add fileName
        end if
    next
    
    set fileList = fList
    set fso = nothing
end function

sub folderExpand(byVal folderName, byRef fList)
    dim fso
    dim objFolder, objSubFolders, objFiles
    set fso = createObject("scripting.filesystemobject")
    set objFolder = fso.getFolder(folderName)
    set objSubFolders = objFolder.subFolders
    set objFiles = objFolder.files
    
    for each objFile in objFiles
        fList.add objFile.path
    next
    
    for each objSubFolder in objSubFolders
        call folderExpand(objSubFolder.path, fList)
    next
    
    set objFiles = nothing
    set objSubFolders = nothing
    set objFolder = nothing
    set fso = nothing
end sub

function fileType(byVal fileName)
    if right(fileName, 1) = "x" then
        fileName = left(fileName, len(fileName) - 1)
    end if
    select case right(filename, 3)
        case "ppt" fileType = 1
        case "doc" fileType = 2
        case else fileType = -1
    end select
end function

function cutFileName(byVal fileName)
    do
        if right(fileName, 1) = "." then
            exit do
        else
            fileName = left(fileName, len(fileName) - 1)
        end if
    loop
    cutFileName = fileName
end function
