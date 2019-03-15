;说明
;  将随笔记(GNnote)中的笔记内容批量复制到OneNote中
;简明流程
;  Loop {
;      GNnote获取文件夹名
;      转到OneNote创建分区
;      Loop {
;          GNnote获取该文件夹下一笔记
;          转到OneNote该分区下创建笔记
;      }
;  }
;
;配置说明
;  1. 将自己的GNote数据库文件复制到电脑端
;       [变量dbPath][other\gnotes是事例数据库, 可用于参考结构]
;       手机端路径[/data/data/org.dayup.gnotes/databases/gnotes]
;       电脑端路径[F:\资料\备份资料\GNote\2019-03-15\gnotes]
;  2. 将自己的GNote附件文件夹复制到电脑端
;       [变量attachmentPath]
;       手机端路径[/storage/emulated/0/.GNotes]
;       电脑端路径[F:\资料\备份资料\GNote\2019-03-15]
;  3. 快捷键[win+P]可以暂停\开启当前脚本 
;
;其他说明
;  1. 本人GNote数据库1.8m, 共计1903条笔记, 24条笔记包含图片附件, 转移笔记花费时长约46分钟
;  2. 除了用户显示自定义创建的文件夹, 任何不在这些文件夹下的笔记会被纳入[GNoteOther]


;========================= 环境配置 =========================
#Persistent
#NoEnv
#SingleInstance, Force
#ErrorStdOut
#HotkeyInterval 1000
SetBatchLines, 10ms
SetTitleMatchMode, 2
SetKeyDelay, -1
StringCaseSense, off
CoordMode, Menu
#Include <JSON> 
#Include <PRINT>
#Include <DBA>
;========================= 环境配置 =========================


;========================= 初始化 =========================
global CurrentDB := Object()
global commonDelay = 400
global commonDelay2 = 4000
global dbPath = "F:\资料\备份资料\GNote\2019-03-15\gnotes"
global attachmentPath = "F:\资料\备份资料\GNote\2019-03-15"
;========================= 初始化 =========================


;========================= 业务逻辑 =========================
print("GNote2OneNote: start.....")
Sleep, %commonDelay2%
DBConnect()
folders := DBFolderFind()
folderOther := Object()
folderOther._id := 0
folderOther.name := "GNoteOther"
folders.push(folderOther)

for folderIndex, folderObj in folders { ; 遍历folder级别    
    BuildOneNoteFolder(folderObj.name)

    notes := DBNotesFindByFolder(folderObj._id)
    for noteIndex, noteObj in notes { ; 遍历note级别
        
        attachments := []
        if (noteObj.is_attach != 0) {
            attachments := DBAttachmentsFindByNote(noteObj._id)
        }
        BuildOneNoteNote(noteIndex, noteObj.content, attachments)
    }
    print("GNote2OneNote: <" folderObj.name "> process over, include <" notes.length() "> notes")
}
print("GNote2OneNote: finish.....")
ExitApp
;========================= 业务逻辑 =========================



;========================= 配置热键 =========================
#P::Pause ;暂停\开启脚本
;========================= 配置热键 =========================







;========================= 公共函数 =========================
BuildOneNoteFolder(folderName) {
    WinActivate,ahk_class Framework::CFrame
    Clipboard := folderName
    Sleep, %commonDelay%
    SendInput, ^t
    Sleep, %commonDelay%
    SendInput, ^v
    Sleep, %commonDelay%
    send {enter}
}
BuildOneNoteNote(noteIndex, noteContent, attachments) {
    WinActivate,ahk_class Framework::CFrame
    if (noteIndex != 1) { ;新分区的第一个页面是onenote自动创建的
        send ^n
        Sleep, %commonDelay%
    }
    
    ;页面标题回车跳过
    send {Enter}
    Sleep, %commonDelay%
    
    ;页面内容
    Clipboard := noteContent
    send ^v
    Sleep, %commonDelay%
    
    
    ;附件图片
    if (attachments && attachments.Length()) {
        send {Enter}
        Sleep, %commonDelay%
        for attachmentIndex, attachmentObj in attachments {
            imgPath := BuildImgPath(attachmentObj.local_path)
            if (!FileExist(imgPath)) ;附件图片不存在时跳过
                continue
            
            CopyImg(imgPath)
            send ^v
            Sleep, %commonDelay2%
            send {Enter}
        }
    }
}

BuildImgPath(imgPath) {
    imgPath := attachmentPath StrReplace(imgPath, "/", "\")
    return imgPath
}
CopyImg(imgPath) {
    RunWait, "other\copyimg.exe" %imgPath%
}
Array2Str(array) {
    if (!array || !array.Length())
        return
    str := ""
    for index, element in array {
        if (index == array.Length())
          str .= element
        else
          str .= element ","
    }
    return str
}
;========================= 公共函数 =========================


;========================= DB-DAO =========================
DBFolderFind() {
    return Query("select _id, name from folder order by _order")
}
DBNotesFindByFolder(folderId) {
    return Query("select reminder_Id, _id, is_attach, content from note  where folder_id = " folderId " ORDER BY modified_time DESC")
}
DBAttachmentsFindByNote(noteId) {
    return Query("select _id, local_path from attachment where note_id = '" noteId "' ORDER BY modified_time DESC")
}
;========================= DB-DAO =========================


;========================= DB-Base =========================
DBConnect() {
	connectionString := dbPath
	try {
		CurrentDB := DBA.DataBaseFactory.OpenDataBase("SQLite", connectionString)
	} catch e
		MsgBox,16, Error, % "Failed to create connection. Check your Connection string and DB Settings!`n`n" ExceptionDetail(e)
}
QueryOne(SQL){
    objs := Query(SQL)
    if (objs.Length())
        return objs[1]
}
Query(SQL){
	if (!IsObject(CurrentDB)) {
        MsgBox, 16, Error, No Connection avaiable. Please connect to a db first!
        return
	}
    SQL := Trim(SQL)
    if (!SQL)
        return
    try {
        resultSet := CurrentDB.OpenRecordSet(SQL)
        if (!is(resultSet, DBA.RecordSet))
            throw Exception("RecordSet Object expected! resultSet was of type: " typeof(resultSet), -1)
        return DBResultSet2Obj(resultSet)
    } catch e {
        MsgBox,16, Error, % "OpenRecordSet Failed.`n`n" ExceptionDetail(e) ;state := "!# " e.What " " e.Message
    }
}

DBResultSet2Obj(resultSet) {
    colNames := resultSet.getColumnNames()
    if (!colNames.Length())
        return Object()
    objs := Object()
    while(!resultSet.EOF){
        obj := Object()
        for index, colName in colNames {
            val := resultSet[colName]
            if (val != DBA.DataBase.NULL)
                obj[colName] := val
        }
        objs.Push(obj)
        resultSet.MoveNext()
    }
    return objs
}
;========================= DB-Base =========================