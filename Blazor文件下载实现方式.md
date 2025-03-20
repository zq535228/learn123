# Blazor 文件下载的几种实现方式

在 Blazor 应用程序中，实现文件下载功能是一个常见需求。本文将介绍几种不同的实现方式，并分析它们的优缺点。

## 1. 使用 JavaScript Interop

这是最简单的实现方式，但需要依赖 JavaScript。

```csharp
public async Task DownloadFile()
{
    await using var memoryStream = new MemoryStream(_processedFile);
    var streamRef = new DotNetStreamReference(stream: memoryStream);
    
    await JS.InvokeVoidAsync("downloadFileFromStream", streamRef, newFileName);
}
```

需要在 JavaScript 中实现对应的方法：

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

## 2. 使用控制器（推荐方式）

这种方式完全不依赖 JavaScript，更加灵活和可控。

首先创建一个下载控制器：

```csharp
[ApiController]
[Route("api/[controller]")]
public class DownloadController : ControllerBase
{
    private readonly ITempFileStorage _tempFileStorage;
    
    public DownloadController(ITempFileStorage tempFileStorage)
    {
        _tempFileStorage = tempFileStorage;
    }
    
    [HttpGet("{fileId}")]
    public async Task<IActionResult> Download(string fileId)
    {
        var (fileContent, fileName) = await _tempFileStorage.GetFileAsync(fileId);
        if (fileContent == null)
        {
            return NotFound();
        }
        
        return File(fileContent, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }
}
```

在 Blazor 页面中使用：

```csharp
<a href="@_downloadUrl" class="btn btn-primary" download="@_downloadFileName">
    <i class="fas fa-download"></i> 下载文件
</a>

@code {
    private string _downloadUrl = string.Empty;
    private string _downloadFileName = string.Empty;

    private async Task ProcessFile()
    {
        // 生成文件ID和下载URL
        var fileId = Guid.NewGuid().ToString("N");
        _downloadUrl = $"/api/download/{fileId}";
        _downloadFileName = "processed_file.xlsx";
        
        // 保存文件到临时存储
        await SaveToTempStorage(fileId, fileContent);
    }
}
```

## 优缺点分析

### JavaScript Interop 方式
优点：
- 实现简单
- 适合小文件下载
- 不需要额外的服务器存储

缺点：
- 依赖 JavaScript
- 不适合大文件
- 难以实现进度显示
- 无法实现复杂的权限控制

### 控制器方式
优点：
- 完全不依赖 JavaScript
- 支持大文件下载
- 可以实现下载进度显示
- 可以实现复杂的权限控制
- 更好的错误处理
- 可以实现断点续传

缺点：
- 需要实现临时文件存储
- 实现相对复杂
- 需要考虑临时文件的清理

## 最佳实践建议

1. 对于简单的小文件下载（<10MB），可以使用 JavaScript Interop 方式
2. 对于大文件或需要更多控制的场景，建议使用控制器方式
3. 实现临时文件存储时，考虑使用 Redis 或文件系统
4. 注意实现定时清理临时文件的机制
5. 在控制器中实现适当的权限验证
6. 考虑添加下载进度显示功能
7. 实现错误处理和重试机制

## 示例代码

可以在 GitHub 上找到完整的示例代码：[Blazor File Download Demo](https://github.com/your-repo)

## 总结

选择合适的文件下载实现方式取决于具体的需求。对于简单场景，JavaScript Interop 方式足够使用；对于复杂场景，建议使用控制器方式以获得更好的可控性和扩展性。无论选择哪种方式，都要注意安全性、性能和用户体验。 

---

> 作者：Claude
> 
> 日期：2024-05-14
> 
> 标签：Blazor, 文件下载, JavaScript, 控制器 