# ahk-GNote2OneNote
> 将随笔记(GNnote)中的笔记内容批量复制到OneNote中


### 简明流程
```
Loop {
    GNnote获取文件夹名
    转到OneNote创建分区
    Loop {
        GNnote获取该文件夹下一笔记
        转到OneNote该分区下创建笔记
    }
}
```

### 配置说明
1. 将自己的GNote数据库文件复制到电脑端
     * [变量dbPath][other\gnotes是事例数据库, 可用于参考结构]
     * 手机端路径[/data/data/org.dayup.gnotes/databases/gnotes]
     * 电脑端路径[F:\资料\备份资料\GNote\2019-03-15\gnotes]
2. 将自己的GNote附件文件夹复制到电脑端
     * [变量attachmentPath]
     * 手机端路径[/storage/emulated/0/.GNotes]
     * 电脑端路径[F:\资料\备份资料\GNote\2019-03-15]
3. 快捷键[win+P]可以暂停\开启当前脚本 


### 其他说明
1. 本人GNote数据库1.8m, 共计1903条笔记, 24条笔记包含图片附件, 转移笔记花费时长约46分钟
2. 除了用户显示自定义创建的文件夹, 任何不在这些文件夹下的笔记会被纳入[GNoteOther]




### 演示
<div align=center><img height="473" width="644" src="https://github.com/bjc5233/ahk-GNote2OneNote/raw/master/resources/demo.gif"/></div>
