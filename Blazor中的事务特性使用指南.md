# Blazor中的事务特性使用指南

在 Blazor 应用程序中，事务管理是确保数据一致性的关键。本文将介绍如何使用 `TransactionalAttribute` 特性来优雅地管理数据库事务。

## 1. 特性介绍

`TransactionalAttribute` 是一个基于 FreeSql 和 Rougamo AOP 框架的特性，用于声明式事务管理。它可以：

1. 自动管理事务的开启、提交和回滚
2. 支持不同的事务传播级别
3. 支持设置事务隔离级别
4. 优雅处理异步方法的事务

## 2. 基本用法

### 2.1 简单事务

```csharp
public class UserService 
{
    [Transactional(Propagation.Required)]
    public async Task CreateUser(UserEntity user)
    {
        await _userRepo.InsertAsync(user);
        await _roleRepo.InsertAsync(user.Role);
        // 如果任何操作失败，整个事务都会回滚
    }
}
```

### 2.2 指定事务级别

```csharp
// 指定传播级别
[Transactional(Propagation.RequiresNew)]
public async Task IndependentOperation() { }

// 指定隔离级别
[Transactional(Propagation.Required, IsolationLevel.ReadCommitted)]
public async Task CriticalOperation() { }
```

## 3. 事务传播级别

1. Required（默认）
   - 使用当前事务，如果没有则创建新事务
   - 适合大多数场景

2. RequiresNew
   - 总是创建新事务
   - 适合独立的操作

3. Nested
   - 在当前事务中创建嵌套事务
   - 适合可以部分回滚的操作

## 4. 使用场景

### 4.1 订单处理

```csharp
[Transactional(Propagation.Required)]
public async Task CreateOrder(OrderEntity order)
{
    // 1. 创建订单
    await _orderRepo.InsertAsync(order);
    
    // 2. 扣减库存
    foreach(var item in order.Items)
    {
        await _stockRepo.DeductStock(item.ProductId, item.Quantity);
    }
    
    // 3. 创建支付记录
    await _paymentRepo.InsertAsync(new PaymentEntity 
    {
        OrderId = order.Id,
        Amount = order.TotalAmount
    });
}
```

### 4.2 批量处理

```csharp
[Transactional(Propagation.Required)]
public async Task ProcessBatchOrders(List<OrderEntity> orders)
{
    foreach(var order in orders)
    {
        await ProcessSingleOrder(order);
    }
}

[Transactional(Propagation.Nested)]
private async Task ProcessSingleOrder(OrderEntity order)
{
    // 处理单个订单
    // 如果失败只回滚这个订单的处理
}
```

## 5. 注意事项

### 5.1 配置要求

在 Blazor 组件中使用事务特性时，必须在 `_Imports.razor` 中注入 `IServiceProvider`：

```razor
@inject IServiceProvider ServiceProvider
```

### 5.2 异常处理

```csharp
[Transactional(Propagation.Required)]
public async Task RiskOperation()
{
    try 
    {
        // 业务操作
    }
    catch (Exception ex)
    {
        // 事务会自动回滚
        // 这里可以处理异常，但不要吞掉异常
        throw;
    }
}
```

### 5.3 性能考虑

1. 避免在小型操作上使用事务
2. 合理设置事务隔离级别
3. 不要在事务中执行耗时的非数据库操作
4. 对于大批量操作，考虑分批处理

## 6. 最佳实践

### 6.1 合理使用事务级别

```csharp
// 简单操作
[Transactional(Propagation.Required)]
public async Task SimpleOperation() { }

// 独立操作
[Transactional(Propagation.RequiresNew)]
public async Task IndependentOperation() { }

// 可嵌套操作
[Transactional(Propagation.Nested)]
public async Task NestedOperation() { }
```

### 6.2 处理大批量数据

```csharp
public async Task ProcessLargeData(List<DataItem> items)
{
    foreach(var batch in items.Chunk(100))
    {
        await ProcessBatch(batch);
    }
}

[Transactional(Propagation.Required)]
private async Task ProcessBatch(DataItem[] batch)
{
    // 处理单个批次
}
```

### 6.3 正确处理异步操作

```csharp
[Transactional(Propagation.Required)]
public async Task AsyncOperation()
{
    // 使用 await 确保事务正确处理
    await Task1();
    await Task2();
    // 避免使用 Task.WhenAll
}
```

## 7. 总结

`TransactionalAttribute` 通过 AOP 的方式优雅地处理了事务管理，使得代码更加清晰和易于维护。合理使用这个特性可以：

1. 提高代码的可读性和可维护性
2. 确保数据的一致性
3. 简化事务管理的复杂度
4. 优雅处理异步操作的事务

在实际开发中，应根据业务场景选择合适的事务传播级别和隔离级别，并注意性能优化。 