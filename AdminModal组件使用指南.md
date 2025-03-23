# AdminModal 组件使用指南

## 组件介绍

AdminModal 是一个基于 Bootstrap 的模态框组件，它提供了以下特性：

1. 可拖拽的模态框
2. 支持全屏显示
3. 支持自定义标题和按钮文本
4. 支持确认和取消事件回调
5. 支持键盘控制
6. 支持静态背景（点击背景不关闭）

## 基本用法

### 1. 最简单的使用方式

```razor
<AdminModal Visible="@showModal" 
            VisibleChanged="@(value => showModal = value)"
            Title="测试模态框">
    <div>这里是模态框的内容</div>
</AdminModal>

@code {
    private bool showModal = false;
}
```

### 2. 带确认按钮的模态框

```razor
<AdminModal Visible="@showModal" 
            VisibleChanged="@(value => showModal = value)"
            Title="确认操作" 
            OnYes="@HandleConfirm">
    <div>确定要执行此操作吗？</div>
</AdminModal>

@code {
    private bool showModal = false;
    
    private async Task HandleConfirm()
    {
        // 处理确认逻辑
        await Task.Delay(1000);
        showModal = false;
    }
}
```

### 3. 自定义按钮文本

```razor
<AdminModal Visible="@showModal" 
            VisibleChanged="@(value => showModal = value)"
            Title="自定义按钮" 
            YesButton="确定" 
            CloseButton="取消">
    <div>自定义按钮文本的模态框</div>
</AdminModal>
```

### 4. 全屏模态框

```razor
<AdminModal Visible="@showModal" 
            VisibleChanged="@(value => showModal = value)"
            Title="全屏模态框" 
            DialogClassName="modal-fullscreen">
    <div>这是一个全屏的模态框</div>
</AdminModal>
```

## 参数说明

| 参数名 | 类型 | 默认值 | 说明 |
|--------|------|--------|------|
| ClientId | string | 自动生成 | 模态框的唯一标识符 |
| IsBackdropStatic | bool | true | 是否启用静态背景（点击背景不关闭） |
| IsKeyboard | bool | true | 是否启用键盘控制（ESC关闭） |
| Title | string | "标题" | 模态框标题 |
| YesButton | string | "保存" | 确认按钮文本 |
| CloseButton | string | "取消" | 取消按钮文本 |
| DialogClassName | string | null | 模态框的CSS类名 |
| Visible | bool | false | 控制模态框显示/隐藏 |
| VisibleChanged | EventCallback<bool> | null | 模态框显示状态改变事件 |
| OnYes | EventCallback | null | 确认按钮点击事件 |
| OnClose | EventCallback | null | 关闭事件（包括点击取消按钮和关闭图标） |

## 高级用法

### 1. 禁用背景点击关闭

```razor
<AdminModal Visible="@showModal" 
            VisibleChanged="@(value => showModal = value)"
            Title="静态背景" 
            IsBackdropStatic="true">
    <div>点击背景不会关闭此模态框</div>
</AdminModal>
```

### 2. 禁用键盘控制

```razor
<AdminModal Visible="@showModal" 
            VisibleChanged="@(value => showModal = value)"
            Title="禁用键盘" 
            IsKeyboard="false">
    <div>按ESC键不会关闭此模态框</div>
</AdminModal>
```

### 3. 自定义样式

```razor
<AdminModal Visible="@showModal" 
            VisibleChanged="@(value => showModal = value)"
            Title="自定义样式" 
            DialogClassName="modal-lg">
    <div>这是一个大尺寸的模态框</div>
</AdminModal>
```

## 注意事项

1. 模态框默认是可拖拽的，通过拖拽标题栏可以移动位置
2. 当使用 `modal-fullscreen` 或 `modal-xxl` 类名时，模态框将不可拖拽
3. 组件内部使用了 jQuery 的 draggable 功能，确保项目中已引入 jQuery
4. 建议在模态框内容较多时使用 `modal-dialog-scrollable` 类名，以支持内容滚动

## 最佳实践

1. 对于简单的提示信息，使用基本的模态框即可
2. 对于表单提交等需要用户确认的操作，使用带确认按钮的模态框
3. 对于复杂的内容展示，考虑使用全屏模态框
4. 在模态框内容较多时，建议添加滚动条支持
5. 对于重要的操作确认，建议禁用背景点击关闭和键盘控制

## 示例代码

以下是一个完整的示例，展示了如何在实际项目中使用 AdminModal：

```razor
@page "/modal-demo"

<Button OnClick="@(() => showModal = true)">打开模态框</Button>

<AdminModal Visible="@showModal" 
            VisibleChanged="@(value => showModal = value)"
            Title="用户信息" 
            YesButton="保存" 
            CloseButton="取消"
            OnYes="@HandleSave"
            OnClose="@HandleClose">
    <div class="p-3">
        <div class="mb-3">
            <label class="form-label">用户名</label>
            <Input @bind-Value="@username" class="form-control" />
        </div>
        <div class="mb-3">
            <label class="form-label">邮箱</label>
            <Input @bind-Value="@email" class="form-control" />
        </div>
    </div>
</AdminModal>

@code {
    private bool showModal = false;
    private string username = "";
    private string email = "";

    private async Task HandleSave()
    {
        // 处理保存逻辑
        await Task.Delay(1000);
        showModal = false;
    }

    private async Task HandleClose()
    {
        // 处理关闭逻辑
        username = "";
        email = "";
    }
}
```

这个示例展示了如何创建一个包含表单的模态框，包括：
- 打开/关闭控制
- 表单数据绑定
- 保存和关闭事件处理
- 基本的表单布局

## 常见问题与解决方案

### 1. 按钮与模态框的交互

以下是一个常见的错误示例：

```razor
<Button onclick="HandleConfirm">handleConfirm</Button>

<AdminModal Visible="@showModal"
            VisibleChanged="@(value => showModal = value)"
            Title="确认操作"
            OnYes="@HandleConfirm">
    <div>确定要执行此操作吗？</div>
</AdminModal>

@code {
    private bool showModal = false;

    private async Task HandleConfirm()
    {
        showModal = !showModal;  // 不推荐这样写
    }
}
```

正确的写法应该是：

```razor
<Button OnClick="@(() => showModal = true)">打开确认框</Button>

<AdminModal Visible="@showModal"
            VisibleChanged="@(value => showModal = value)"
            Title="确认操作"
            OnYes="@HandleConfirm">
    <div>确定要执行此操作吗？</div>
</AdminModal>

@code {
    private bool showModal = false;

    private async Task HandleConfirm()
    {
        // 处理确认逻辑
        await Task.Delay(1000);  // 模拟异步操作
        showModal = false;       // 确认后关闭模态框
    }
}
```

主要改进点：
1. 按钮的 `onclick` 改为 `OnClick`（Blazor 中事件处理使用 Pascal 命名）
2. 按钮点击时直接设置 `showModal = true` 来打开模态框
3. 在 `HandleConfirm` 中处理完业务逻辑后，设置 `showModal = false` 来关闭模态框
4. 避免在同一个方法中切换模态框的显示状态，这可能导致用户体验不佳

### 2. 模态框状态管理的最佳实践

1. **打开模态框**：
   ```razor
   <Button OnClick="@(() => showModal = true)">打开模态框</Button>
   ```

2. **关闭模态框**：
   ```razor
   // 在确认按钮事件中
   private async Task HandleConfirm()
   {
       // 处理业务逻辑
       await Task.Delay(1000);
       showModal = false;
   }

   // 在取消按钮事件中
   private async Task HandleCancel()
   {
       // 清理数据
       showModal = false;
   }
   ```

3. **使用 OnClose 事件**：
   ```razor
   <AdminModal Visible="@showModal"
               VisibleChanged="@(value => showModal = value)"
               Title="确认操作"
               OnYes="@HandleConfirm"
               OnClose="@HandleClose">
       <div>确定要执行此操作吗？</div>
   </AdminModal>

   @code {
       private async Task HandleClose()
       {
           // 处理关闭逻辑，如清理表单数据
           await Task.Delay(100);
           showModal = false;
       }
   }
   ``` 