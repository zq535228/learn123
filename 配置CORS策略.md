# 在ASP.NET Core中配置CORS策略

在现代Web应用程序中，跨域资源共享（CORS）是一个重要的安全特性。它允许或限制不同来源的Web应用程序访问服务器资源。在ASP.NET Core中，我们可以通过配置CORS策略来控制这些访问。

## 配置步骤

### 1. 添加CORS服务

首先，在 `Program.cs` 文件中，我们需要添加CORS服务。我们可以通过 `AddCors` 方法来定义一个CORS策略。在这个例子中，我们定义了一个名为 `"CorsPolicy"` 的策略：

```csharp
builder.Services.AddCors(options => {
    options.AddPolicy("CorsPolicy", builder => {
        builder
            .WithOrigins("http://localhost:3000")
            .WithOrigins("https://ai.jianyandashu.com")
            .AllowAnyMethod()
            .AllowAnyHeader()
            .AllowCredentials()
            .SetIsOriginAllowed(_ => true);  // 临时添加，用于调试
    });
});
```

- **AllowAnyMethod**: 允许任何HTTP方法（GET, POST, PUT, DELETE等）。
- **AllowAnyHeader**: 允许任何请求头。
- **AllowCredentials**: 允许请求携带凭证（如Cookies）。
- **SetIsOriginAllowed**: 这里使用了 `SetIsOriginAllowed(_ => true)`，这将允许所有来源的请求。这是一个临时的调试设置，建议在生产环境中使用更严格的配置。

### 2. 使用CORS中间件

在构建应用程序后，我们需要在中间件管道中使用定义的CORS策略：

```csharp
var app = builder.Build();
app.UseCors("CorsPolicy");
```

这行代码确保了所有的HTTP请求都会应用我们定义的CORS策略。

## 注意事项

- **安全性**: 在生产环境中，建议根据实际需求限制允许的来源，而不是使用 `SetIsOriginAllowed(_ => true)`。
- **策略名称**: 确保在 `UseCors` 中使用的策略名称与 `AddCors` 中定义的名称一致。

通过以上步骤，我们可以在ASP.NET Core应用程序中有效地配置CORS策略，确保应用程序的安全性和灵活性。 