# SelectTable2组件使用指南

SelectTable2 是一个基于 Blazor 和 FreeSql 的表格选择组件，它直接在页面上展示可选择的数据表格，适用于需要即时选择数据的场景。

## 基本特性

1. 组件参数
- `TItem`: 数据实体类型(必须是 class 且可实例化)
- `TKey`: 主键类型
- `Value`: 单选时绑定的值
- `Items`: 多选时绑定的列表
- `PageSize`: 分页大小(默认20)
- `IsSearchText`: 是否启用搜索功能(默认true)
- `ChildContent`: 自定义内容模板
- `OnQuery`: 查询数据的回调方法

2. 主要功能
- 支持单选和多选模式
- 内置分页功能
- 支持文本搜索
- 自动处理 FreeSql 实体关系
- 支持自定义显示模板

## 使用示例

### 1. 单选模式
```razor
<SelectTable2 TItem="User" TKey="long" @bind-Value="selectedUserId">
    <ChildContent Context="user">
        @user.Name (@user.Department)
    </ChildContent>
</SelectTable2>
```

### 2. 多选模式
```razor
<SelectTable2 TItem="Role" TKey="int" @bind-Items="selectedRoles">
    <ChildContent Context="role">
        @role.Name - @role.Description
    </ChildContent>
</SelectTable2>
```

### 3. 带查询条件的示例
```razor
<SelectTable2 TItem="Product" TKey="Guid" 
    @bind-Items="selectedProducts"
    OnQuery="HandleQuery">
    <ChildContent Context="product">
        @product.Name (@product.Price.ToString("C"))
    </ChildContent>
</SelectTable2>

@code {
    private async Task HandleQuery(AdminQueryEventArgs<Product> args)
    {
        // 添加查询条件
        args.Select.Where(a => a.IsActive == true)
            .OrderByDescending(a => a.CreateTime);
    }
}
```

## 注意事项

1. 实体类要求：
   - 必须是 class 类型
   - 必须有无参构造函数
   - 必须有单一主键
   - 主键类型必须与 TKey 类型匹配

2. 性能优化：
   - 建议设置合适的 PageSize
   - 在 OnQuery 中添加必要的查询条件
   - 避免加载不必要的数据

3. 使用建议：
   - 对于简单的选择场景，优先使用 SelectTable2
   - 需要弹窗选择时，使用 InputTable2
   - 自定义显示内容时，善用 ChildContent 模板

## 与 InputTable2 的区别

1. 展现形式：
   - SelectTable2：直接在页面上显示表格
   - InputTable2：点击按钮后弹窗显示表格

2. 使用场景：
   - SelectTable2：适合直接在页面上选择数据
   - InputTable2：适合需要弹窗选择或者表单中选择数据

3. 功能差异：
   - SelectTable2：功能相对简单，界面更清爽
   - InputTable2：功能更丰富，支持更多自定义选项

## 最佳实践

1. 简单选择场景：
```razor
<SelectTable2 TItem="Category" TKey="int" @bind-Value="categoryId">
    <ChildContent Context="category">
        @category.Name
    </ChildContent>
</SelectTable2>
```

2. 带格式化的选择：
```razor
<SelectTable2 TItem="Employee" TKey="long" @bind-Items="selectedEmployees">
    <ChildContent Context="emp">
        <div class="d-flex align-items-center">
            <img src="@emp.Avatar" class="rounded-circle" width="32" />
            <span class="ml-2">@emp.Name</span>
            <small class="text-muted ml-2">(@emp.Department)</small>
        </div>
    </ChildContent>
</SelectTable2>
```

3. 结合查询条件：
```razor
<SelectTable2 TItem="Order" TKey="Guid" 
    @bind-Items="selectedOrders"
    IsSearchText="true"
    PageSize="10"
    OnQuery="args => args.Select
        .Where(a => a.Status == OrderStatus.Active)
        .OrderByDescending(a => a.CreateTime)">
    <ChildContent Context="order">
        订单号：@order.OrderNo - @order.Amount.ToString("C")
    </ChildContent>
</SelectTable2>
``` 