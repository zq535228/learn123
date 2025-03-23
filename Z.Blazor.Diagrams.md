# Z.Blazor.Diagrams：强大的Blazor图表组件库

## 1. 简介

Z.Blazor.Diagrams（简称ZBD）是一个全面、可定制、可扩展的通用图表库，专为Blazor（服务器端和WebAssembly）开发的。它最初受到React库react-diagrams的启发，但后来发展成为一个功能更加丰富的解决方案。ZBD可用于创建具有自定义设计的高级图表，甚至库的行为也是"可黑客化"的，可以根据特定需求进行调整。

## 2. 核心特性

- **多用途**：适用于各种图表需求，如流程图、网络拓扑图、数据库设计等
- **触控支持**：完全支持触控设备操作
- **双层结构**：使用SVG层绘制连线，HTML层绘制节点，实现最大化的自定义能力
- **全面的连线功能**：支持节点间、端口间甚至连线间的连接
- **丰富的布局工具**：连线路由器、路径生成器、标记和标签
- **交互功能**：支持平移、缩放和自适应缩放
- **选择功能**：多选、删除和区域选择
- **分组支持**：将组作为一等公民，具有节点的所有功能
- **自定义组件**：支持自定义节点、连线和组
- **可替换行为**：可自定义链接拖动、模型删除等行为
- **导航器**：可定制的图表概览/导航器，适用于大型图表
- **网格对齐**：支持对齐到网格
- **虚拟化**：只渲染用户可见的节点，提高性能
- **锁定机制**：支持只读模式
- **算法支持**：提供各种图表算法

## 3. 安装方法

要在项目中使用ZBD，需要安装以下NuGet包：

```csharp
<PackageReference Include="Z.Blazor.Diagrams" Version="3.0.2" />
<PackageReference Include="Z.Blazor.Diagrams.Core" Version="3.0.2" />
<PackageReference Include="Z.Blazor.Diagrams.Algorithms" Version="3.0.2" />
```

## 4. 基本使用示例

### 4.1 基础组件设置

在Razor页面中添加DiagramCanvas控件：

```html
<div class="wrapper">
    <Content>
        <CascadingValue Value="BlazorDiagram">
            <div class="diagram-container">
                <DiagramCanvas>
                    <Widgets>
                        <NavigatorWidget Width="200"
                                         Height="150"
                                         Style="position: absolute; bottom: 15px; right: 15px;"></NavigatorWidget>
                    </Widgets>
                </DiagramCanvas>
            </div>
        </CascadingValue>
    </Content>
</div>
```

### 4.2 初始化和配置

在代码中初始化图表：

```csharp
@code {
    private Diagram BlazorDiagram { get; set; }

    protected override void OnInitialized()
    {
        // 创建图表实例
        BlazorDiagram = new Diagram(new DiagramOptions
        {
            DeleteKey = "Delete",
            AllowMultiSelection = true,
            AllowPanning = true,
            GridSize = 15,
            Zoom = new DiagramZoomOptions
            {
                Enabled = true,
                Inverse = true,
                Min = 0.1,
                Max = 2
            }
        });

        // 添加节点
        var node1 = new NodeModel(new Point(50, 50));
        var node2 = new NodeModel(new Point(300, 150));
        BlazorDiagram.AddNode(node1);
        BlazorDiagram.AddNode(node2);

        // 添加连线
        BlazorDiagram.AddLink(new LinkModel(node1, node2));
    }
}
```

### 4.3 自定义节点

创建自定义节点组件：

```html
@inherits NodeWidget

<g>
    <rect width="@Node.Size.Width" 
          height="@Node.Size.Height" 
          fill="lightblue" 
          stroke="blue" 
          stroke-width="2"/>
    <text x="@(Node.Size.Width / 2)" 
          y="@(Node.Size.Height / 2)" 
          text-anchor="middle" 
          alignment-baseline="middle">
        @((Node as CustomNodeModel).Title)
    </text>
</g>
```

并在C#代码中定义自定义节点模型：

```csharp
public class CustomNodeModel : NodeModel
{
    public string Title { get; set; }

    public CustomNodeModel(Point position, string title) : base(position)
    {
        Title = title;
        Size = new Size(100, 50);
    }
}
```

然后在初始化代码中注册：

```csharp
BlazorDiagram.RegisterModelComponent<CustomNodeModel, CustomNodeComponent>();
```

## 5. 高级功能

### 5.1 端口（Ports）

端口是节点上的连接点，可以让连线更加精确地连接到节点的特定位置：

```csharp
// 创建带端口的节点
var nodeWithPorts = new NodeModel(new Point(200, 200));
// 添加端口
var topPort = new PortModel(PortAlignment.Top);
var bottomPort = new PortModel(PortAlignment.Bottom);
nodeWithPorts.AddPort(topPort);
nodeWithPorts.AddPort(bottomPort);
BlazorDiagram.AddNode(nodeWithPorts);
```

### 5.2 分组

分组功能允许将多个节点组合在一起：

```csharp
// 创建组
var group = new GroupModel(new Point(100, 100), 300, 200);
BlazorDiagram.AddGroup(group);

// 将节点添加到组
var nodeInGroup = new NodeModel(new Point(150, 150));
BlazorDiagram.AddNode(nodeInGroup);
group.AddNode(nodeInGroup);
```

### 5.3 事件处理

ZBD提供丰富的事件处理能力：

```csharp
protected override void OnInitialized()
{
    // 初始化...

    // 注册事件
    BlazorDiagram.SelectionChanged += OnSelectionChanged;
    BlazorDiagram.LinkAdded += OnLinkAdded;
}

private void OnSelectionChanged(SelectionChangedEventArgs args)
{
    // 处理选择变更事件
    Console.WriteLine($"选中了 {args.Selected.Count()} 个项目");
}

private void OnLinkAdded(LinkModel link)
{
    // 处理添加连线事件
    Console.WriteLine("添加了新连线");
}
```

## 6. 实际应用场景

ZBD适用于多种应用场景，包括但不限于：

- **流程图**：构建业务流程、工作流程图
- **网络拓扑图**：可视化网络结构和连接
- **数据库设计**：创建实体关系图和数据库模型
- **组织结构图**：展示公司或团队的层级结构
- **状态机图**：可视化系统状态和转换
- **自定义图表**：任何需要节点和连线的可视化需求

## 7. 性能优化

对于大型图表，ZBD提供了虚拟化功能，只渲染用户可见的节点，以提高性能：

```csharp
BlazorDiagram = new Diagram(new DiagramOptions
{
    // 其他选项...
    VirtualizationEnabled = true,
    VirtualizationOptions = new DiagramVirtualizationOptions
    {
        RenderNodeBuffer = 200 // 可见区域外渲染缓冲
    }
});
```

## 8. 总结

Z.Blazor.Diagrams是一个强大而灵活的Blazor图表组件库，为开发者提供了构建复杂图表应用的完整解决方案。通过其丰富的特性和可定制性，开发者可以快速创建满足特定需求的交互式图表。

无论是简单的流程图还是复杂的网络拓扑图，ZBD都能胜任，是Blazor开发者的得力助手。随着对性能的持续优化和功能的不断扩展，ZBD有望在Blazor生态系统中占据更重要的位置。

## 9. 相关资源

- GitHub仓库：[https://github.com/Blazor-Diagrams/Blazor.Diagrams](https://github.com/Blazor-Diagrams/Blazor.Diagrams)
- 官方文档：[https://blazor-diagrams.zhaytam.com/](https://blazor-diagrams.zhaytam.com/)
- 示例项目：[https://github.com/Blazor-Diagrams/Blazor.DatabaseDesigner](https://github.com/Blazor-Diagrams/Blazor.DatabaseDesigner)
