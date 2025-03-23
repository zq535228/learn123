# AdminTable2组件使用指南

AdminTable2是一个基于Blazor和FreeSql的通用表格组件,提供了丰富的表格展示和操作功能。本文将详细介绍其使用方法。

## 基本用法

```razor
<AdminTable2 TItem="User">
    <TableHeader>
        <th>用户名</th>
        <th>年龄</th>
        <th>创建时间</th>
    </TableHeader>
    <TableRow Context="user">
        <td>@user.Username</td>
        <td>@user.Age</td>
        <td>@user.CreateTime</td>
    </TableRow>
</AdminTable2>
```

## 主要特性

### 1. 基础配置

- `Title`: 弹框标题
- `PageSize`: 分页大小,默认20
- `IsQueryString`: 是否将查询条件同步到URL
- `BodyHeight`: 表格主体高度,默认-1(自适应)

### 2. 功能开关

- `IsRemove`: 开启删除功能
- `IsRowRemove`: 开启行内删除按钮
- `IsAdd`: 开启添加功能
- `IsEdit`: 开启编辑功能
- `IsRefersh`: 开启刷新功能
- `IsSearchText`: 开启文本搜索
- `IsSingleSelect`: 开启单选
- `IsMultiSelect`: 开启多选
- `IsAutoSelectParent`: 自动选中父节点
- `IsConfirmEdit`: 编辑时确认弹框
- `IsConfirmRemove`: 删除时确认弹框

### 3. 布局模式

- `Colspan`: 一行显示几条记录,默认4
- `TableTd99Width`: 操作列宽度,默认160

### 4. 模板定制

- `TableHeader`: 表头模板
- `TableRow`: 行模板
- `TableTh1`: 第一列表头模板
- `TableTd1`: 第一列内容模板
- `TableTd99`: 操作列模板
- `EditTemplate`: 添加/编辑弹框模板
- `CardHeader`: 卡片头部模板
- `CardFooter`: 卡片底部模板

### 5. 事件回调

- `InitQuery`: 初始化查询条件
- `OnQuery`: 查询时触发
- `OnEdit`: 编辑时触发
- `OnEditFinish`: 编辑完成时触发
- `OnRemove`: 删除时触发
- `OnSelectChanged`: 选择变化时触发
- `OnRowClick`: 点击行时触发

## 高级用法

### 1. 树形表格

组件会自动识别实体中的树形关系(OneToMany),并支持树形展示:

```csharp
public class Menu
{
    public int Id { get; set; }
    public string Name { get; set; }
    public int? ParentId { get; set; }
    public List<Menu> Children { get; set; }
}
```

### 2. 自定义查询

```csharp
<AdminTable2 TItem="User" OnQuery="@OnQuery">
    ...
</AdminTable2>

@code {
    async Task OnQuery(AdminQueryEventArgs<User> args)
    {
        args.Select = args.Select.Where(a => a.Age > 18);
    }
}
```

### 3. 自定义编辑模板

```razor
<AdminTable2 TItem="User" EditTemplate="@EditTemplate">
    ...
</AdminTable2>

@code {
    RenderFragment<User> EditTemplate = user => @<div>
        <div class="form-group">
            <label>用户名</label>
            <input @bind="user.Username" class="form-control" />
        </div>
        <div class="form-group">
            <label>年龄</label>
            <input @bind="user.Age" class="form-control" />
        </div>
    </div>;
}
```

### 4. 行内操作按钮

```razor
<AdminTable2 TItem="User" TableTd99="@TableTd99">
    ...
</AdminTable2>

@code {
    RenderFragment<User> TableTd99 = user => @<div>
        <button @onclick="@(() => HandleCustomAction(user))" class="btn btn-sm btn-info">
            自定义操作
        </button>
    </div>;
}
```

## 注意事项

1. 实体类必须使用单一主键
2. 编辑时会自动处理实体间的关联关系(OneToOne/OneToMany/ManyToOne)
3. 树形表格模式下会自动禁用分页
4. 建议在编辑模板中使用 `@bind` 而不是 `@bind:event="oninput"` 以避免频繁刷新

## 最佳实践

1. 合理使用模板定制,避免过度自定义
2. 对于复杂查询,使用 `OnQuery` 事件进行扩展
3. 编辑模板建议使用表单验证
4. 大量数据时注意设置 `BodyHeight` 启用滚动
5. 合理使用 `IsConfirmEdit` 和 `IsConfirmRemove` 防止误操作 