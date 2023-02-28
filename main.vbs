option explicit
dim objShell, objFile, objExcel
set objShell = CreateObject("wscript.shell")
set objFile = CreateObject("Scripting.FileSystemObject")
set objExcel = CreateObject("Excel.Application") 
dim directory, posX, posY, resX, resY, posZ, configData, mapTrees, mapPos0, mapPos1, mapLogoff, mapDots, mapNote1, mapLock1, mapHack

sub Main()
    directory = objFile.GetParentFolderName(wscript.ScriptFullName)
    'Debug()
    Initialize()
    Render()

    do
        CheckInput()
        Render()
        wscript.Sleep(100)
        objShell.SendKeys("{F5}")
    loop
end sub

sub Initialize()
    configData = objFile.OpenTextFile(directory &  "\config.txt").ReadAll()
    resX = GetKeyValue(configData, "resX")
    resX = cint(resX)
    resY = GetKeyValue(configData, "resY")
    resY = cint(resY)
    posX = GetKeyValue(configData, "posX")
    posX = cint(posX)
    posY = GetKeyValue(configData, "posY")
    posY = cint(posY)
    posZ = GetKeyValue(configData, "posZ")
    posZ = cint(posZ)
    LoadMap()
    ClearFolder()
end sub

sub LoadMap()
    dim map

    select case posZ
    case 0
        map = "pos0"
    case 1
        map = "pos1"
    end select

    mapTrees = GetKeyValue(configData, map)
    mapPos1 = GetKeyValue(configData, map)
    mapPos0 = GetKeyValue(configData, map)
    mapLogoff = GetKeyValue(configData, map)
    mapDots = GetKeyValue(configData, map)
    mapNote1 = GetKeyValue(configData, map)
    mapLock1 = GetKeyValue(configData, map)
    mapHack = GetKeyValue(configData, map)
    mapTrees = ParseMap(mapTrees, "trees")
    mapPos1 = ParseMap(mapPos1, "pos1")
    mapPos0 = ParseMap(mapPos0, "pos0")
    mapLogoff = ParseMap(mapLogoff, "logoff")
    mapDots = ParseMap(mapDots, "dots")
    mapNote1 = ParseMap(mapNote1, "note1")
    mapLock1 = ParseMap(mapLock1, "lock1")
    mapHack = ParseMap(mapHack, "hack")
end sub

sub Render()
    dim total, i, pos
    total = resX * resY
    pos = (resX * posY) + posX

    for i = 0 to total - 1
        if i = pos then
            call objFile.CopyFile(directory & "\assets\player.lnk", directory & "\game\" & i & ".lnk", true)
        elseif ArrayContainsValue(mapTrees, i) then
            call objFile.CopyFile(directory & "\assets\tree.lnk", directory & "\game\" & i & ".lnk", true)
        elseif ArrayContainsValue(mapLogoff, i) then
            call objFile.CopyFile(directory & "\assets\logoff.lnk", directory & "\game\" & i & ".lnk", true)
        elseif ArrayContainsValue(mapDots, i) then
            call objFile.CopyFile(directory & "\assets\dot.lnk", directory & "\game\" & i & ".lnk", true)
        elseif ArrayContainsValue(mapNote1, i) then
            call objFile.CopyFile(directory & "\assets\note.lnk", directory & "\game\" & i & ".lnk", true)
        elseif ArrayContainsValue(mapLock1, i) then
            call objFile.CopyFile(directory & "\assets\lock.lnk", directory & "\game\" & i & ".lnk", true)
        elseif ArrayContainsValue(mapHack, i) then
            call objFile.CopyFile(directory & "\assets\hack.lnk", directory & "\game\" & i & ".lnk", true)
        else
            call objFile.CopyFile(directory & "\assets\blank.lnk", directory & "\game\" & i & ".lnk", true)
        end if
    next
end sub

function CheckCollision(checkX, checkY)
    dim pos
    pos = (resX * checkY) + checkX

    if checkX < 0 then
        CheckCollision = false
        exit function
    end if

    if checkX >= resX then
        CheckCollision = false
        exit function
    end if

    if checkY < 0 then
        CheckCollision = false
        exit function
    end if

    if checkY >= resY then
        CheckCollision = false
        exit function
    end if

    if ArrayContainsValue(mapTrees, pos) then
        CheckCollision = false
        exit function
    end if

    if ArrayContainsValue(mapLogoff, pos) then
        CheckCollision = false
        objShell.Run("logoff")
        exit function
    end if

    if ArrayContainsValue(mapLock1, pos) then
        CheckCollision = false
        exit function
    end if

    if ArrayContainsValue(mapNote1, pos) then
        call objShell.Run("""" & directory & "\assets\files\note1.txt""",, true)
        CheckCollision = false
        exit function
    end if

    if ArrayContainsValue(mapHack, pos) then
        call objShell.Run("""" & directory & "\assets\files\hack.bat""", 3, true)
        CheckCollision = false
        exit function
    end if

    if ArrayContainsValue(mapPos1, pos) then
        CheckCollision = true
        posX = checkX
        posY = 1
        posZ = 1
        LoadMap()
        exit function
    end if

    if ArrayContainsValue(mapPos0, pos) then
        CheckCollision = true
        posX = checkX
        posY = 5
        posZ = 0
        LoadMap()
        exit function
    end if

    posX = checkX
    posY = checkY
    CheckCollision = true
end function

sub ClearFolder()
    dim files, element
    set files = objFile.GetFolder(directory & "\game").Files

    for each element in files
        objFile.DeleteFile(element.Path)
    next
end sub

sub CheckInput()
    dim res

    do
        if isKeyPressed(87) then
            if CheckCollision(posX, posY - 1) then
                exit sub
            end if
        end if

        if isKeyPressed(65) then
            if CheckCollision(posX - 1, posY) then
                exit sub
            end if
        end if

        if isKeyPressed(83) then
            if CheckCollision(posX, posY + 1) then
                exit sub
            end if
        end if

        if isKeyPressed(68) then
            if CheckCollision(posX + 1, posY) then
                exit sub
            end if
        end if

        if isKeyPressed(27) then
            res = msgbox("Are you sure you want to quit?", 32 + 4)

            if res = 6 then
                ClearFolder()
                wscript.Quit
            end if
        end if

        TerminateShortcut()
        wscript.Sleep(100)
    loop
end sub

function ParseMap(mapData, mapType)
    dim element, i, startPos
    dim temp
    mapData = split(mapData, ",")

    for i = 0 to ubound(mapData)
        mapData(i) = trim(mapData(i))
    next

    for i = 0 to ubound(mapData)
        if mapData(i) = mapType then
            startPos = i + 1
            exit for
        end if
    next

    for i = startPos to ubound(mapData)
        if mapData(i) = "end" then
            temp = left(temp, len(temp) - 2)
            temp = split(temp, vbcrlf)
            exit for
        else
            temp = temp + mapData(i) + vbcrlf
        end if
    next

    for i = 0 to ubound(temp)
        temp(i) = eval(temp(i))
    next

    ParseMap = temp
end function

function PushArray(inputArray, push)
    dim length, i
    length = ubound(inputArray)
    execute("dim temp(" & length + 1 & ")")

    for i = 0 to length
        temp(i) = inputArray(i)
    next

    temp(length + 1) = push
    PushArray = temp
end function

function GetKeyValue(haystack, needle)
    dim element, key, config
    config = split(haystack, vbcrlf)

    for each element in config
        element = trim(element)
        key = left(element, instr(element, ":") - 1)
        key = trim(key)

        if key = needle then
            GetKeyValue = mid(element, instr(element, ":") + 1)
            GetKeyValue = trim(GetKeyValue)
            exit function
        end if
    next
end function

function isKeyPressed(keyValue)
    dim api, cmd
    api = 0
    cmd = "CALL(""user32.dll"", ""GetAsyncKeyState"", ""JJ"", " & keyValue & ")"
    api = objExcel.ExecuteExcel4Macro(cmd)

    if api <> 0 then
        isKeyPressed = true
    else
        isKeyPressed = false
    end if
end function

'Checks if shift(16), alt(18), and T(84) keys are pressed
sub TerminateShortcut()
    if isKeyPressed(16) then
        if isKeyPressed(18) then
            if isKeyPressed(84) then
                wscript.Echo("The program has been terminated")
                wscript.Quit
            end if
        end if
    end if
end sub

function ArrayContainsValue(haystack, needle)
    dim element

    for each element in haystack
        if trim(element) = trim(needle) then
            ArrayContainsValue = true
            exit function
        end if
    next

    ArrayContainsValue = false
end function

sub Breakpoint(message)
    dim toString, element

    if isarray(message) then
        for each element in message
            toString = toString & element & vbcrlf
        next

        message = toString
    end if

    wscript.Echo(message)
    wscript.Quit
end sub

sub Debug()
    dim test
    test = array("The", "quick", "brown")
    test = PushArray(test, "fox")
    Breakpoint(test)
end sub

Main()