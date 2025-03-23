# AllocTable2组件使用指南

AllocTable2是一个基于Blazor和FreeSql的数据分配组件，主要用于处理一对多、多对多等关系的数据分配场景。本文将详细介绍其使用方法。

## 组件特点

1. 基于FreeSql的实体关系
2. 支持一对多、多对多数据分配
3. 自动处理数据状态
4. 提供丰富的自定义模板
5. 支持树形结构展示

## 基本用法

```razor
<AllocTable2 TItem="RoleEntity" TChild="UserEntity" 
    @bind-Item="@allocItemUsers" 
    ChildProperty="Users" 
    Title="@($"【分配用户】{allocItemUsers?.Name}")">
    <TableTd1>@context.Username</TableTd1>
</AllocTable2>

@code {
    RoleEntity allocItemUsers;
    
    void BeginAllocUsers(RoleEntity item)
    {
        allocItemUsers = item;
    }
}
```

## 主要参数

### 1. 泛型参数
- `TItem`: 主实体类型
- `TChild`: 子实体类型

### 2. 基础属性
- `Item`: 当前正在分配的主实体对象
- `ChildProperty`: 主实体中的子集合属性名称
- `Title`: 弹窗标题
- `PageSize`: 分页大小，默认20条
- `IsSearchText`: 是否启用搜索功能，默认true

### 3. 模板参数
- `TableTh1`: 第一列表头模板
- `TableTd1`: 第一列内容模板
- `TableHeader`: 表格头部模板
- `TableRow`: 表格行模板

### 4. 事件回调
- `ItemChanged`: Item属性变化时的回调
- `OnQuery`: 查询数据时的回调

## 使用场景

### 1. 角色分配用户

```razor
<AllocTable2 TItem="RoleEntity" TChild="UserEntity" 
    @bind-Item="@allocItemRoles" 
    ChildProperty="Users" 
    Title="@($"【分配用户】{allocItemRoles?.Name}")">
    <TableTd1>@context.Username</TableTd1>
</AllocTable2>
```

### 2. 角色分配菜单

```razor
<AllocTable2 TItem="RoleEntity" TChild="MenuEntity" 
    @bind-Item="@allocItemMenus" 
    ChildProperty="Menus" 
    Title="@($"【分配菜单】{allocItemMenus?.Name}")" 
    OnQuery="BeginAllocMenusOnQuery">
    <TableTd1>@context.Label</TableTd1>
</AllocTable2>

@code {
    void BeginAllocMenusOnQuery(AdminQueryEventArgs<MenuEntity> e)
    {
        var menuIds = admin.RoleMenus.Select(b => b.Id);
        e.Select.WhereIf(admin.Tenant.Id != "main", 
            a => menuIds.Contains(a.Id))
            .OrderBy(a => a.Sort);
    }
}
```

## 工作原理

1. 组件会自动处理一对多、多对多关系的数据绑定
2. 使用FreeSql的导航属性功能自动加载关联数据
3. 支持树形结构的数据展示
4. 提供选择、重置、确认等基本操作
5. 自动维护数据状态和关系

## 注意事项

1. 实体类必须使用单一主键
2. 需要正确配置FreeSql的实体关系映射
3. `ChildProperty`必须是主实体中存在的导航属性
4. 建议使用`@bind-Item`双向绑定来处理数据更新

## 最佳实践

1. 合理使用`OnQuery`事件来过滤和排序数据
2. 使用`TableTd1`来自定义每行的显示内容
3. 根据需要设置合适的`PageSize`来优化性能
4. 使用`Title`属性来提供清晰的操作提示

## 总结

AllocTable2组件极大地简化了数据分配的操作流程，通过与FreeSql的深度集成，使得复杂的数据关系处理变得简单直观。合理使用这个组件，可以快速实现角色权限分配、用户组管理等常见的后台管理功能。 