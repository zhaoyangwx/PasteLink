# PasteLink
右键粘贴符号连接

注意事项：若创建软链接，请勿删除源文件。如需防误删请粘贴硬链接

功能说明：

粘贴为硬连接：仅限同一分区。读取剪贴板内容，如果是目录则递归遍历，按照原目录结构创建一个相同的目录结构，并对遍历到的所有文件按照原先的结构单独创建链接。如果是文件则执行mklink /h创建硬链接。

粘贴为符号连接：读取剪贴板内容，如果是目录则执行mklink /j命令，创建指向该目录的符号连接。如果是文件则执行mklink，对文件创建符号连接。

复制目录，粘贴文件为符号连接：读取剪贴板内容，如果是目录则递归遍历，按照原目录结构创建一个相同的目录结构，并对遍历到的所有文件按照原先的结构单独创建链接。如果是文件则创建符号链接。

注意：源目录内部不能有递归结构（比如存在符号连接指向其父路径或自身）。

PS.复制时可以多选



