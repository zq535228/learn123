# Blazor中使用MiniExcel处理Excel文件

在Blazor开发中，经常会遇到需要处理Excel文件的场景。本文将介绍如何使用MiniExcel这个轻量级库来实现Excel的读写操作，并结合一个实际的条码加密工具案例来说明。

## MiniExcel简介

MiniExcel是一个轻量级的.NET Excel读写库，它的主要特点是：

1. 低内存占用：采用流式处理方式
2. 使用简单：API设计直观
3. 支持大文件处理：适合在Web应用中使用
4. 支持.NET Core：完全跨平台

## 实际案例：条码加密工具

### 功能描述

这是一个用于处理Excel文件中条码加密的工具，主要功能包括：
- 上传Excel文件
- 读取并处理文件中的条码
- 将加密后的数据保存为新的Excel文件
- 提供文件下载功能

### 代码实现

1. 首先引入必要的命名空间：

```csharp
@using MiniExcelLibs
@using System.IO
```

2. Excel文件读取：

```csharp
// 读取Excel数据
using var stream = new MemoryStream();
await file.OpenReadStream().CopyToAsync(stream);
stream.Position = 0;
            
var rows = await stream.QueryAsync();
var dataList = rows.ToList();
```

3. 数据处理：

```csharp
// 处理每一行的条码
foreach (IDictionary<string, object> row in dataList)
{
    var barcode = row["B"]?.ToString();
    if (!string.IsNullOrEmpty(barcode))
    {
        var encryptedBarcode = await BarcodeEncryption.EncryptAsync(barcode);
        row["B"] = encryptedBarcode;
        _processedCount++;
    }
}
```

4. 保存处理后的文件：

```csharp
using var outputStream = new MemoryStream();
await outputStream.SaveAsAsync(dataList);
_processedFile = outputStream.ToArray();
```

5. 文件下载实现：

```csharp
// JavaScript 互操作函数
private async Task DownloadFile()
{
    try
    {
        if (_processedFile == null) return;

        var fileName = Path.GetFileNameWithoutExtension(_fileName);
        var extension = Path.GetExtension(_fileName);
        var newFileName = $"{fileName}_encrypted{extension}";

        // 使用 Blazor 的文件下载功能
        using var memoryStream = new MemoryStream(_processedFile);
        using var streamRef = new DotNetStreamReference(stream: memoryStream);
        
        await JS.InvokeVoidAsync("downloadFileFromStream", streamRef, newFileName);
    }
    catch (Exception ex)
    {
        Logger.LogError(ex, "下载文件时发生错误");
        throw;
    }
}
```

6. JavaScript 文件下载处理：

```javascript
window.downloadFileFromStream = async (streamRef, fileName) => {
    const arrayBuffer = await streamRef.arrayBuffer();
    const blob = new Blob([arrayBuffer]);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
} 
```

## 技术要点

1. **流式处理**
   - 使用MemoryStream处理文件上传
   - MiniExcel的流式读写减少内存占用

2. **动态数据处理**
   - 使用IDictionary<string, object>处理Excel行数据
   - 支持动态列访问

3. **异步操作**
   - 文件读写采用异步方法
   - 确保Web应用的响应性

4. **用户体验**
   - 添加加载状态显示
   - 处理完成后显示结果统计
   - 提供直观的下载按钮

## 注意事项

1. 错误处理
   - 需要对文件读写操作进行异常处理
   - 记录错误日志便于问题排查

2. 内存管理
   - 使用using语句确保资源正确释放
   - 及时清理临时文件和流

3. 安全性
   - 验证上传文件的类型
   - 控制文件大小限制

## 总结

MiniExcel在Blazor应用中处理Excel文件是一个很好的选择，它提供了简单易用的API，同时保持了较低的内存占用。通过本文的案例，我们可以看到它在处理文件上传、数据处理和文件下载等场景中的应用。

## 参考资料

- [MiniExcel官方文档](https://github.com/shps951023/MiniExcel)
- [Blazor文件处理最佳实践](https://docs.microsoft.com/aspnet/core/blazor/file-uploads) 