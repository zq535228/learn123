# Blazor 权限菜单特性使用指南

## 前言

在 Blazor 后台管理系统中，权限管理是一个重要的功能。通过使用 `AdminButton` 特性，我们可以方便地为页面添加权限按钮，这些按钮会自动出现在后台的权限分配界面中。本文将详细介绍如何使用 `AdminButton` 特性来实现权限管理。

## AdminButton 特性介绍

`AdminButton` 特性是一个用于标记方法的特性，它可以：
1. 自动生成权限菜单项
2. 在后台权限分配界面中显示
3. 支持动态权限控制

## 基本用法

### 1. 简单示例

```razor
@page "/demo-page"

<Button OnClick="_CreateInferredCallback">创建按钮</Button>

@code {
    private bool showModal = false;

    [AdminButton("TestAdminButton")]
    private void _CreateInferredCallback()
    {
        // 按钮点击处理逻辑
    }
}
```

### 2. 完整示例

```razor
@page "/user-management"

<div class="container">
    <Button OnClick="_CreateUser">创建用户</Button>
    <Button OnClick="_DeleteUser">删除用户</Button>
    <Button OnClick="_EditUser">编辑用户</Button>
</div>

@code {
    [AdminButton("CreateUser")]
    private void _CreateUser()
    {
        // 创建用户逻辑
    }

    [AdminButton("DeleteUser")]
    private void _DeleteUser()
    {
        // 删除用户逻辑
    }

    [AdminButton("EditUser")]
    private void _EditUser()
    {
        // 编辑用户逻辑
    }
}
```

## 工作原理

1. **特性标记**
   - 使用 `[AdminButton]` 特性标记需要权限控制的方法
   - 特性参数为权限标识符，在权限系统中唯一标识此操作

2. **权限注册**
   - 系统启动时会扫描所有带有 `AdminButton` 特性的方法
   - 将这些方法注册到权限系统中

3. **权限分配**
   - 在后台管理界面中可以看到所有注册的权限项
   - 管理员可以为不同角色分配不同的权限

## 使用步骤

1. **添加特性**
   ```razor
   [AdminButton("UniqueButtonName")]
   private void _MethodName()
   {
       // 方法实现
   }
   ```

2. **重新登录后台**
   - 添加新的 AdminButton 特性后需要重新登录后台
   - 新添加的权限项会出现在权限分配界面

3. **分配权限**
   - 在后台权限管理界面中为角色分配权限
   - 可以精确控制每个按钮的访问权限

## 最佳实践

1. **命名规范**
   - 权限标识符使用有意义的名称
   - 建议使用模块名+操作名的形式，如：`UserManagement_Create`

2. **方法命名**
   - 建议在方法名前加下划线，如：`_CreateUser`
   - 这样可以清晰地区分带权限控制的方法

3. **权限粒度**
   - 根据实际需求合理划分权限粒度
   - 避免权限划分过细或过粗

4. **权限组织**
   - 相关的权限项放在同一个页面中
   - 便于后期维护和管理

## 注意事项

1. **特性参数**
   - 特性参数必须唯一
   - 建议使用有意义的名称，方便权限管理

2. **重新登录**
   - 添加新的权限项后必须重新登录后台
   - 否则新的权限项不会出现在权限分配界面

3. **权限验证**
   - 系统会自动验证用户是否有权限执行相应操作
   - 无权限的按钮会被自动隐藏或禁用

## 实际应用示例

### 1. 用户管理页面
```razor
@page "/user-management"

<div class="container">
    <div class="row mb-3">
        <div class="col">
            <Button OnClick="_CreateUser">新增用户</Button>
            <Button OnClick="_BatchDelete">批量删除</Button>
        </div>
    </div>
    
    <Table>
        <!-- 表格内容 -->
    </Table>
</div>

@code {
    [AdminButton("UserManagement_Create")]
    private void _CreateUser()
    {
        // 创建用户逻辑
    }

    [AdminButton("UserManagement_BatchDelete")]
    private void _BatchDelete()
    {
        // 批量删除逻辑
    }
}
```

### 2. 系统设置页面
```razor
@page "/system-settings"

<div class="container">
    <Button OnClick="_SaveSettings">保存设置</Button>
    <Button OnClick="_ResetSettings">重置设置</Button>
</div>

@code {
    [AdminButton("SystemSettings_Save")]
    private void _SaveSettings()
    {
        // 保存设置逻辑
    }

    [AdminButton("SystemSettings_Reset")]
    private void _ResetSettings()
    {
        // 重置设置逻辑
    }
}
```

## 总结

`AdminButton`  
1. 轻松实现权限控制
2. 自动生成权限配置项
3. 提高系统的可维护性

在实际开发中，建议遵循最佳实践，合理规划权限结构，这样可以让系统的权限管理更加清晰和高效。 