# FreeSql和Blazor实现后台权限管理系统

## 前言
在开发企业级后台管理系统时,权限管理是一个基础且重要的功能。本文将介绍如何使用 FreeSql + Blazor 实现一个功能完善的后台权限管理系统。

## 核心功能设计

### 1. 多租户管理
系统支持多租户架构,通过域名自动识别不同租户。每个租户可以拥有独立的数据库,实现数据完全隔离。

```csharp
public IFreeSql Orm
{
    get
    {
        if (_orm != null) return _orm;
        // 通过域名识别租户
        var tenantHost = HttpContext.Request.Host.Host.ToLower();
        Tenant = cloud.Use("main").Select<TenantEntity>()
            .Where(a => a.Host == tenantHost && a.IsEnabled)
            .First();
        
        if (Tenant == null) return _orm = cloud.Use("main");
        return _orm = GetTenantFreeSql(Tenant.Id);
    }
}
```

### 2. 用户认证
实现了基于 Cookie 的用户认证机制,支持记住登录状态。同时还包含了防止多端同时登录的逻辑。

```csharp
// 用户登录
async public Task SignIn(UserEntity user, bool remember)
{
    user.LoginTime = DateTime.Now;
    await Orm.Update<UserEntity>()
        .Where(a => a.Id == user.Id)
        .Set(a => a.LoginTime, user.LoginTime)
        .ExecuteAffrowsAsync();

    // 设置登录Cookie
    HttpContext.Response.Cookies.Append("login", 
        DesEncrypt.Encrypt(user.Id.ToString() + "|" + user.LoginTime.ToString("yyyy-MM-dd HH:mm:ss")), 
        new CookieOptions
        {
            Path = "/",
            Expires = remember ? DateTimeOffset.UtcNow.AddDays(15) : null
        });
}
```

### 3. 权限管理
系统实现了基于角色的访问控制(RBAC),支持页面级和按钮级的权限控制。

```csharp
// 页面权限验证
async public Task<bool> AuthPath(string path)
{
    path = path?.ToLower().Trim('/');
    if (User == null) return false;
    await InitRoles();
    CurrentMenu = RoleMenus.Where(a => a.PathLower == path).FirstOrDefault();
    return AuthPathSuccess = CurrentMenu != null;
}

// 按钮权限验证
async public Task<bool> AuthButton(string path)
{
    if (User == null || CurrentMenu == null) return false;
    await InitRoles();
    // ... 按钮权限验证逻辑
}
```

### 4. 数据库自动同步
利用 FreeSql 的自动同步特性,系统会自动维护数据库结构。

```csharp
internal static void ConfigFreeSql(IFreeSql fsql)
{
    // 配置数据库时间
    var serverTime = fsql.Ado.QuerySingle(() => DateTime.UtcNow);
    var timeOffset = DateTime.UtcNow.Subtract(serverTime);
    
    // 配置实体审计
    fsql.Aop.AuditValue += (_, e) =>
    {
        // ... 审计逻辑
    };
}
```

## 使用方法

### 1. 服务注册
```csharp
services.AddScoped<AdminContext>();
```

### 2. Blazor组件中使用
```csharp
@inject AdminContext AdminContext

protected override async Task OnInitializedAsync()
{
    await AdminContext.Init();
    
    if (!await AdminContext.AuthPath("/admin/users"))
    {
        AdminContext.RedirectLogin();
        return;
    }
}
```

### 3. 权限控制示例
```csharp
@if (await AdminContext.AuthButton("add"))
{
    <Button OnClick="@AddUser">添加用户</Button>
}
```

## 实现要点

1. **多租户数据隔离**
   - 每个租户使用独立的数据库
   - 通过域名自动识别租户
   - 支持租户级别的权限配置

2. **用户会话管理**
   - 使用加密Cookie存储会话信息
   - 支持记住登录状态
   - 防止多端同时登录

3. **权限验证机制**
   - 支持页面级权限控制
   - 支持按钮级权限控制
   - 支持动态权限配置

4. **性能优化**
   - 使用缓存减少数据库查询
   - 支持连接池
   - 异步操作提高响应速度

## 注意事项

1. 在多租户场景下要特别注意数据隔离
2. 合理设计菜单和按钮的权限结构
3. 注意处理并发登录的情况
4. 根据业务需求扩展权限验证逻辑
5. 定期清理过期的会话信息

## 总结

通过 FreeSql + Blazor 的组合,我们可以快速构建一个功能完善的后台权限管理系统。系统不仅支持基础的用户认证和权限控制,还实现了多租户管理等企业级功能。这个方案可以作为开发企业管理系统的良好起点。 