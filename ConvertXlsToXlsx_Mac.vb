Sub ConvertXlsToXlsx_Mac()
    Dim sourceFolderPath As String
    Dim destinationFolderPath As String
    Dim fileName As String
    Dim sourceFilePath As String
    Dim destinationFilePath As String
    Dim xlApp As Object
    Dim xlWorkbook As Object
    
    ' 设置源文件夹和目标文件夹的路径
    ' 注意：在macOS上，路径格式通常是 POSIX 风格的 "/Users/YourUsername/FolderName"
    ' 确保这些路径与您的系统相匹配
    sourceFolderPath = "/Users/hang/Library/Containers/com.tencent.xinWeChat/Data/Library/Application Support/com.tencent.xinWeChat/2.0b4.0.9/237e6ba0b2b7af52a974319f16a23be7/Message/MessageTemp/c1c30910da9401121316d3c0d92ec2ed/File" 
    destinationFolderPath = "/Users/hang/Library/Containers/com.tencent.xinWeChat/Data/Library/Application Support/com.tencent.xinWeChat/2.0b4.0.9/237e6ba0b2b7af52a974319f16a23be7/Message/MessageTemp/c1c30910da9401121316d3c0d92ec2ed/File" 
    
    ' 如果目标文件夹不存在，则创建它（可选）
    ' 在VBA中，没有直接创建文件夹的函数，但可以使用Shell命令
    ' 如果您的Excel for Mac不支持Shell命令，请手动创建文件夹
    ' If Dir(destinationFolderPath, vbDirectory) = "" Then
    '     MkDir destinationFolderPath
    ' End If
    
    ' 初始化Excel应用程序对象（对于在Excel内部运行的宏，这通常不是必需的）
    ' 但在某些情况下（如在其他Office应用程序中调用Excel），这是必需的
    ' 由于我们是在Excel中运行此宏，因此可以省略此步骤，除非您需要在独立的应用程序中调用Excel
    ' Set xlApp = CreateObject("Excel.Application")
    
    ' 获取源文件夹中的第一个.xls文件
    fileName = Dir(sourceFolderPath & "/*.xls")
    
    ' 遍历源文件夹中的所有.xls文件
    Application.ScreenUpdating = False ' 关闭屏幕更新以提高性能
    Application.DisplayAlerts = False ' 关闭自动弹出的提示信息
    On Error Resume Next ' 忽略错误，以防某个文件无法转换
    Do While fileName <> ""
        ' 构建完整的源文件路径
        sourceFilePath = sourceFolderPath & "/" & fileName
        
        ' 构建目标文件路径，只更改扩展名
        destinationFilePath = destinationFolderPath & "/" & Left(fileName, InStrRev(fileName, ".") - 1) & ".xlsx"
        
        ' 打开.xls文件
        ' 注意：在macOS上，如果Excel for Mac的VBA不支持直接打开文件，
        ' 您可能需要使用AppleScript或其他方法来打开文件并进行转换
        ' 这里我们假设可以直接使用Workbooks.Open
        Set xlWorkbook = Workbooks.Open(sourceFilePath)
        
        ' 另存为.xlsx格式的文件
        xlWorkbook.SaveAs fileName:=destinationFilePath, FileFormat:=xlOpenXMLWorkbook ' xlOpenXMLWorkbook 对应于 .xlsx 格式
        
        ' 关闭工作簿，不保存更改（因为我们已经通过SaveAs保存了新的.xlsx文件）
        xlWorkbook.Close SaveChanges:=False
        
        ' 释放对象变量
        Set xlWorkbook = Nothing
        
        ' 获取下一个.xls文件
        fileName = Dir
    Loop
    On Error GoTo 0 ' 恢复正常的错误处理
    Application.ScreenUpdating = True ' 恢复屏幕更新
    Application.DisplayAlerts = True ' 恢复自动弹出的提示信息
    
    ' 如果之前创建了xlApp对象，则应该在这里关闭它（对于本宏，这不是必需的）
    ' xlApp.Quit
    ' Set xlApp = Nothing
    
    MsgBox "所有.xls文件已成功转换为.xlsx文件，并保存在 " & destinationFolderPath, vbInformation
End Sub