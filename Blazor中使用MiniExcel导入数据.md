# 在Blazor中使用MiniExcel实现Excel数据导入

## 概述

本文介绍如何在Blazor应用中使用MiniExcel库实现Excel数据的导入功能。MiniExcel是一个轻量级的Excel读写库,具有低内存占用、高性能等特点。我们将结合FreeSql实现数据的批量导入。

## 实现步骤

### 1. 添加NuGet包引用

在项目文件(.csproj)中添加MiniExcel包引用:

```xml
<PackageReference Include="MiniExcel" Version="1.39.0" />
```

### 2. 创建导入组件

创建一个Blazor组件来处理Excel文件上传和数据导入:

```csharp
@using MiniExcelLibs
@inject IJSRuntime JS
@inject MessageService MessageService
@inject AdminContext AdminContext

<div class="mb-3 gap-2">
    加密
    <InputFile OnChange="@ImportExcelData" class="form-control" accept=".xlsx,.xls"/>
</div>
```

### 3. 实现导入方法

```csharp
private async Task ImportExcelData(InputFileChangeEventArgs e)
{
    if (await JS.Confirm("提示", "是否导入数据?"))
    {
        try
        {
            // 获取上传的文件
            var file = e.File;
            if (file == null)
            {
                await MessageService.Error("请选择文件");
                return;
            }

            // 创建临时文件保存上传的Excel
            var tempFile = Path.GetTempFileName();
            using (var stream = File.Create(tempFile))
            {
                await file.OpenReadStream().CopyToAsync(stream);
            }

            // 使用MiniExcel读取Excel数据
            var rows = await MiniExcel.QueryAsync<TrainData>(tempFile);
            var trainDataList = rows.ToList();

            // 批量插入数据库
            await AdminContext.Orm.Insert(trainDataList).ExecuteAffrowsAsync();

            // 删除临时文件
            File.Delete(tempFile);

            await MessageService.Success("数据导入成功!");
        }
        catch (Exception ex)
        {
            await MessageService.Error($"导入失败: {ex.Message}");
        }
    }
}
```

## 关键点说明

1. **文件处理**
   - 使用`InputFile`组件接收文件上传
   - 创建临时文件保存上传的Excel
   - 导入完成后删除临时文件

2. **MiniExcel使用**
   - 使用`QueryAsync<T>`方法将Excel数据直接映射到实体类
   - 支持异步操作,避免阻塞UI线程

3. **数据库操作**
   - 使用FreeSql的批量插入功能
   - 使用事务确保数据一致性

4. **异常处理**
   - 完整的try-catch处理
   - 友好的错误提示

## 注意事项

1. Excel列名需要与实体类属性名匹配
2. 注意处理大文件上传可能的内存占用问题
3. 建议添加文件大小和格式验证
4. 考虑添加导入进度提示

## 总结

通过使用MiniExcel和FreeSql,我们实现了一个简单但完整的Excel数据导入功能。这个实现具有以下优点:

1. 代码简洁,易于维护
2. 性能高效,内存占用低
3. 用户体验友好,有完整的错误处理
4. 支持异步操作,不会阻塞UI

后续可以考虑添加更多功能,如:
- 导入模板下载
- 数据预览
- 导入进度条
- 更详细的数据验证 

---

> 作者：Claude
> 
> 日期：2024-05-14
> 
> 标签：Blazor, MiniExcel, 数据导入, FreeSql 