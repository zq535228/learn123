# DES加密解密工具使用指南

`DesEncrypt` 是一个基于 DES 算法的加密解密工具类，提供了字符串的加密和解密功能。它支持两种编码方式：Base64和16进制。

## 1. 基本用法

### 1.1 默认加密（DES+Base64）

```csharp
// 使用默认密钥加密
string encrypted = DesEncrypt.Encrypt("需要加密的内容");

// 使用自定义密钥加密（密钥必须是8位）
string encrypted = DesEncrypt.Encrypt("需要加密的内容", "12345678");
```

### 1.2 默认解密（DES+Base64）

```csharp
// 使用默认密钥解密
string decrypted = DesEncrypt.Decrypt(encrypted);

// 使用自定义密钥解密（密钥必须与加密时相同）
string decrypted = DesEncrypt.Decrypt(encrypted, "12345678");
```

## 2. 16进制编码

### 2.1 16进制加密

```csharp
// 默认大写
string encrypted = DesEncrypt.Encrypt4Hex("需要加密的内容");

// 指定小写
string encrypted = DesEncrypt.Encrypt4Hex("需要加密的内容", null, true);

// 自定义密钥
string encrypted = DesEncrypt.Encrypt4Hex("需要加密的内容", "12345678");
```

### 2.2 16进制解密

```csharp
// 使用默认密钥解密
string decrypted = DesEncrypt.Decrypt4Hex(encrypted);

// 使用自定义密钥解密
string decrypted = DesEncrypt.Decrypt4Hex(encrypted, "12345678");
```

## 3. 实际应用场景

### 3.1 密码加密

```csharp
public class UserService
{
    public void CreateUser(string username, string password)
    {
        var user = new UserEntity
        {
            Username = username,
            // 存储加密后的密码
            Password = DesEncrypt.Encrypt(password)
        };
        // 保存用户
    }

    public bool ValidatePassword(string password, string encryptedPassword)
    {
        // 验证密码
        return DesEncrypt.Decrypt(encryptedPassword) == password;
    }
}
```

### 3.2 敏感数据传输

```csharp
public class DataTransferService
{
    private const string TransferKey = "DataKey88";

    public string EncryptData(object data)
    {
        var json = System.Text.Json.JsonSerializer.Serialize(data);
        return DesEncrypt.Encrypt(json, TransferKey);
    }

    public T DecryptData<T>(string encryptedData)
    {
        var json = DesEncrypt.Decrypt(encryptedData, TransferKey);
        return System.Text.Json.JsonSerializer.Deserialize<T>(json);
    }
}
```

### 3.3 配置信息加密

```csharp
public class ConfigService
{
    public void SaveConfig(string key, string value)
    {
        // 加密配置值
        var encrypted = DesEncrypt.Encrypt4Hex(value);
        // 保存到配置文件或数据库
    }

    public string GetConfig(string key)
    {
        // 获取加密的配置值
        var encrypted = "从存储中获取的加密值";
        // 解密
        return DesEncrypt.Decrypt4Hex(encrypted);
    }
}
```

## 4. 注意事项

1. 密钥长度要求：
   - 密钥必须是8位字符
   - 如果不指定密钥，将使用默认密钥 "freesql!"

2. 空值处理：
   - 如果加密/解密的字符串为null或空，将返回null

3. 异常处理：
```csharp
try
{
    var encrypted = DesEncrypt.Encrypt("测试", "123");  // 密钥少于8位
}
catch (ArgumentException ex)
{
    // 将抛出异常：秘钥长度为8位
}
```

4. 编码选择：
   - Base64编码：标准的编码方式，输出字符串较短
   - 16进制编码：可读性更好，适合调试和日志记录

## 5. 最佳实践

1. 密钥管理
```csharp
public static class AppKeys
{
    // 集中管理密钥
    public const string UserPasswordKey = "User@Key8";
    public const string ConfigKey = "Conf@Key8";
    public const string TransferKey = "Data@Key8";
}
```

2. 包装常用方法
```csharp
public static class SecurityHelper
{
    public static string EncryptPassword(string password)
    {
        return DesEncrypt.Encrypt(password, AppKeys.UserPasswordKey);
    }

    public static bool ValidatePassword(string password, string encrypted)
    {
        return DesEncrypt.Decrypt(encrypted, AppKeys.UserPasswordKey) == password;
    }
}
```

3. 异常处理
```csharp
public static string SafeEncrypt(string content, string key = null)
{
    try
    {
        return DesEncrypt.Encrypt(content, key);
    }
    catch (Exception ex)
    {
        // 记录日志
        return null;
    }
}
```

## 6. 总结

`DesEncrypt` 工具类提供了简单易用的 DES 加密解密功能：

1. 支持Base64和16进制两种编码方式
2. 可以使用自定义密钥
3. 适合处理敏感数据的加密需求
4. 集成在项目中使用方便

在实际使用中，建议：

1. 统一管理密钥
2. 根据场景选择合适的编码方式
3. 做好异常处理
4. 考虑将常用操作封装成辅助方法 