# Blazor 模态框尺寸与样式指南

## 前言

在 Blazor 开发中，模态框（Modal）是一个常用的 UI 组件。通过设置不同的 `DialogClassName`，我们可以实现不同的展示效果。本文将详细介绍 Bootstrap 模态框的各种尺寸和样式类名的使用方法。

## 基础尺寸类

Bootstrap 提供了多种预定义的模态框尺寸，可以通过 `DialogClassName` 来设置：

| 类名 | 宽度 | 使用场景 | 示例 |
|------|------|----------|------|
| `modal-sm` | 300px | 简单的提示信息、确认框 | 删除确认框 |
| `modal-lg` | 800px | 表单、详细信息展示 | 用户编辑表单 |
| `modal-xl` | 1140px | 复杂表格、多列表单 | 数据统计表格 |
| `modal-xxl` | 1400px | 大型数据展示、图表 | 数据分析面板 |
| `modal-fullscreen` | 100% | 全屏展示 | 图片预览、复杂工作台 |

### 示例代码

```razor
// 小型模态框
<AdminModal Visible="@showModal" 
            VisibleChanged="@(value => showModal = value)"
            DialogClassName="modal-sm"
            Title="删除确认">
    <div>确定要删除这条记录吗？</div>
</AdminModal>

// 大型模态框
<AdminModal Visible="@showModal" 
            VisibleChanged="@(value => showModal = value)"
            DialogClassName="modal-lg"
            Title="用户信息编辑">
    <div>
        <!-- 用户表单内容 -->
    </div>
</AdminModal>

// 全屏模态框
<AdminModal Visible="@showModal" 
            VisibleChanged="@(value => showModal = value)"
            DialogClassName="modal-fullscreen"
            Title="数据分析面板">
    <div>
        <!-- 复杂的数据展示内容 -->
    </div>
</AdminModal>
```

## 响应式全屏类

Bootstrap 5 引入了响应式全屏模态框，可以根据不同的屏幕尺寸自动调整显示方式：

| 类名 | 全屏触发点 | 使用场景 |
|------|------------|----------|
| `modal-fullscreen-sm-down` | <576px | 手机端全屏 |
| `modal-fullscreen-md-down` | <768px | 平板竖屏全屏 |
| `modal-fullscreen-lg-down` | <992px | 平板横屏全屏 |
| `modal-fullscreen-xl-down` | <1200px | 小型笔记本全屏 |
| `modal-fullscreen-xxl-down` | <1400px | 大型显示器全屏 |

### 示例代码

```razor
// 在手机端自动全屏的模态框
<AdminModal Visible="@showModal" 
            VisibleChanged="@(value => showModal = value)"
            DialogClassName="modal-lg modal-fullscreen-sm-down"
            Title="响应式表单">
    <div>
        <!-- 表单内容 -->
    </div>
</AdminModal>
```

## 辅助样式类

除了尺寸类，Bootstrap 还提供了一些辅助样式类来增强模态框的展示效果：

### 1. 滚动控制
- `modal-dialog-scrollable`：启用模态框内容区域的滚动条
```razor
<AdminModal DialogClassName="modal-lg modal-dialog-scrollable">
    <!-- 长内容 -->
</AdminModal>
```

### 2. 位置控制
- `modal-dialog-centered`：模态框垂直居中显示
```razor
<AdminModal DialogClassName="modal-dialog-centered">
    <!-- 内容 -->
</AdminModal>
```

### 3. 组合使用
```razor
<AdminModal DialogClassName="modal-lg modal-dialog-scrollable modal-dialog-centered">
    <!-- 内容 -->
</AdminModal>
```

## 最佳实践建议

1. **尺寸选择**
   - 简单提示信息：使用默认尺寸或 `modal-sm`
   - 表单内容：使用 `modal-lg`
   - 数据展示：使用 `modal-xl` 或 `modal-xxl`
   - 复杂工作台：使用 `modal-fullscreen`

2. **响应式设计**
   - 移动端优先：使用 `modal-fullscreen-sm-down`
   - 内容复杂：组合使用响应式类和滚动类

3. **性能优化**
   - 全屏模态框慎用，可能影响性能
   - 大量数据展示时，考虑分页或虚拟滚动

4. **用户体验**
   - 内容较多时添加 `modal-dialog-scrollable`
   - 需要突出显示时使用 `modal-dialog-centered`
   - 表单类建议使用固定尺寸，避免频繁改变大小

## 注意事项

1. **拖拽功能**
   - 使用 `modal-fullscreen` 或 `modal-xxl` 时，拖拽功能会自动禁用
   - 可以通过自定义 JS 代码重新启用拖拽

2. **样式冲突**
   - 多个类名组合使用时注意优先级
   - 自定义样式可能会覆盖默认行为

3. **移动端适配**
   - 在移动端，建议使用响应式全屏类
   - 注意控制内容密度，确保可读性

## 实际应用示例

### 1. 图片预览模态框
```razor
<AdminModal DialogClassName="modal-fullscreen modal-dialog-centered"
            Title="图片预览">
    <div class="text-center">
        <img src="@imageUrl" class="img-fluid" />
    </div>
</AdminModal>
```

### 2. 响应式表单模态框
```razor
<AdminModal DialogClassName="modal-lg modal-fullscreen-sm-down modal-dialog-scrollable"
            Title="用户信息编辑">
    <div class="p-3">
        <form>
            <!-- 表单内容 -->
        </form>
    </div>
</AdminModal>
```

### 3. 数据分析面板
```razor
<AdminModal DialogClassName="modal-xxl modal-dialog-scrollable"
            Title="数据分析">
    <div class="row">
        <!-- 复杂的数据展示和图表 -->
    </div>
</AdminModal>
```

## 总结

合理使用 `DialogClassName` 可以让我们的模态框更加灵活和美观。在实际开发中，应根据具体需求选择合适的类名组合，同时注意移动端适配和性能优化。通过本文的介绍，相信你已经掌握了 Bootstrap 模态框的各种尺寸和样式类名的使用方法。 