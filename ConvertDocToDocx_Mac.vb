Sub ConvertDocToDocx_Mac()
    Dim sourceFolderPath As String
    Dim destinationFolderPath As String
    Dim fileName As String
    Dim sourceFilePath As String
    Dim destinationFilePath As String
    Dim doc As Document
    
    ' 设置源文件夹和目标文件夹的路径
    ' 注意：在macOS上，路径格式通常是 POSIX 风格的 "/Users/YourUsername/FolderName"
    sourceFolderPath = "/Users/hang/Library/Containers/com.tencent.xinWeChat/Data/Library/Application Support/com.tencent.xinWeChat/2.0b4.0.9/237e6ba0b2b7af52a974319f16a23be7/Message/MessageTemp/c1c30910da9401121316d3c0d92ec2ed/File" 
    destinationFolderPath = "/Users/hang/Library/Containers/com.tencent.xinWeChat/Data/Library/Application Support/com.tencent.xinWeChat/2.0b4.0.9/237e6ba0b2b7af52a974319f16a23be7/Message/MessageTemp/c1c30910da9401121316d3c0d92ec2ed/File" 
    
    ' 获取源文件夹中的第一个文件
    fileName = Dir(sourceFolderPath & "/*.doc")
    
    ' 遍历源文件夹中的所有.doc文件
    Application.ScreenUpdating = False ' 关闭屏幕更新以提高性能
    Application.DisplayAlerts = False ' 关闭自动弹出的提示信息
    On Error Resume Next ' 忽略错误，以防某个文件无法转换
    Do While fileName <> ""
        ' 构建完整的源文件路径
        sourceFilePath = sourceFolderPath & "/" & fileName
        
        ' 构建目标文件路径，只更改扩展名
        destinationFilePath = destinationFolderPath & "/" & Left(fileName, InStrRev(fileName, ".") - 1) & ".docx"
        
        ' 打开.doc文件
        Set doc = Documents.Open(sourceFilePath)
        
        ' 另存为.docx格式的文件
        doc.SaveAs2 fileName:=destinationFilePath, FileFormat:=wdFormatXMLDocument
        
        ' 关闭文档，不保存更改（因为我们已经通过SaveAs2保存了新的.docx文件）
        doc.Close SaveChanges:=False
        
        ' 释放对象变量
        Set doc = Nothing
        
        ' 获取下一个.doc文件
        fileName = Dir
    Loop
    On Error GoTo 0 ' 恢复正常的错误处理
    Application.ScreenUpdating = True ' 恢复屏幕更新
    Application.DisplayAlerts = True ' 恢复自动弹出的提示信息
    
    MsgBox "所有.doc文件已成功转换为.docx文件，并保存在 " & destinationFolderPath, vbInformation
End Sub

