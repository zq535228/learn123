# C# MD5加密工具类使用说明

MD5(Message-Digest Algorithm 5)是一种被广泛使用的密码散列函数，可以产生出一个128位（16字节）的散列值（hash value），用于确保信息传输完整一致。本文介绍一个实用的MD5加密工具类的使用方法。

## 功能特点

1. 支持16位、32位、64位三种MD5加密方式
2. 支持文件MD5值获取
3. 支持大小写输出控制
4. 支持null值处理

## 代码实现

```csharp
public class Md5Encrypt
{
    /// <summary>
    /// 16位MD5加密
    /// </summary>
    public static string Encrypt16(string password, bool lowerCase = false)
    {
        if (password.IsNull())
            return null;

        using var md5 = MD5.Create();
        return md5.ComputeHash(Encoding.UTF8.GetBytes(password)).ToHex(lowerCase);
    }

    /// <summary>
    /// 32位MD5加密
    /// </summary>
    public static string Encrypt32(string password = "", bool lowerCase = false)
    {
        if (password.IsNull())
            return null;

        using var md5 = MD5.Create();
        string pwd = string.Empty;
        byte[] s = md5.ComputeHash(Encoding.UTF8.GetBytes(password));
        var format = lowerCase ? "x2" : "X2";
        foreach (var item in s)
        {
            pwd = string.Concat(pwd, item.ToString(format));
        }
        return pwd;
    }

    /// <summary>
    /// 64位MD5加密
    /// </summary>
    public static string Encrypt64(string password)
    {
        if (password.IsNull())
            return null;

        using var md5 = MD5.Create();
        byte[] s = md5.ComputeHash(Encoding.UTF8.GetBytes(password));
        return Convert.ToBase64String(s);
    }

    /// <summary>
    /// 获取文件MD5值
    /// </summary>
    public static string GetHash(Stream stream)
    {
        StringBuilder sb = new();
        using var md5 = MD5.Create();
        byte[] hashBytes = md5.ComputeHash(stream);
        foreach (byte bt in hashBytes)
        {
            sb.Append(bt.ToString("x2"));
        }
        return sb.ToString();
    }
}
```

## 使用示例

```csharp
// 32位大写MD5
string pwd1 = Md5Encrypt.Encrypt32("123456"); 
// 结果：E10ADC3949BA59ABBE56E057F20F883E

// 32位小写MD5
string pwd2 = Md5Encrypt.Encrypt32("123456", true); 
// 结果：e10adc3949ba59abbe56e057f20f883e

// 16位MD5
string pwd3 = Md5Encrypt.Encrypt16("123456");
// 结果：49BA59ABBE56E057

// 64位MD5(Base64)
string pwd4 = Md5Encrypt.Encrypt64("123456");
// 结果：4QrcOUm6Wau+VuBX8g+IPg==

// 获取文件MD5
using (var stream = File.OpenRead("文件路径"))
{
    string md5 = Md5Encrypt.GetHash(stream);
}
```

## 使用场景

1. **密码加密存储**：在用户注册或修改密码时，将密码进行MD5加密后存储
2. **文件完整性校验**：通过比对文件的MD5值，确认文件是否被修改
3. **数字签名**：在某些场景下用作简单的数字签名

## 注意事项

1. MD5是不可逆的加密算法，无法从MD5值反推原始数据
2. 为了提高安全性，建议在密码加密时配合加盐（Salt）使用
3. 如果对安全性要求极高的场景，建议使用更安全的加密算法（如SHA256、bcrypt等）
4. 该工具类已对null值进行了处理，使用时无需额外判断

## 总结

本文介绍的MD5加密工具类提供了多种加密方式，使用简单方便，适合大多数常见的MD5加密场景。在实际使用中，建议根据具体的安全需求，选择合适的加密方式和额外的安全措施。 