# FindYourSeat - 智能座位安排系统

一个专为企业培训、会议活动设计的智能座位分配工具，支持角色均匀分布、可视化预览、拖拽调整和多格式导出。

## ✨ 特性

- 🚀 **智能分配算法** - 自动实现角色均匀分布，避免同类角色聚集
- 📊 **Excel数据支持** - 支持上传Excel文件，智能识别员工信息
- 🎯 **灵活配置** - 可自定义会议室数量、桌子数量和座位数
- 🎨 **可视化预览** - 直观的座位图表展示，支持实时预览
- 🖱️ **拖拽调整** - 支持手动拖拽调整座位安排
- 📋 **多格式导出** - 支持Excel、CSV格式导出，按桌分组显示
- 💾 **配置管理** - 支持配置保存、加载和分享
- 📱 **响应式设计** - 适配不同屏幕尺寸，外企风格界面

## 🖥️ 在线使用

本系统为纯前端静态网页，可直接在浏览器中运行。推荐部署到 GitHub Pages 或类似静态网站托管服务。

### GitHub Pages 部署

1. Fork 本仓库
2. 在仓库设置中启用 GitHub Pages
3. 选择 `main` 分支作为源
4. 访问 `https://yourusername.github.io/findyourseat`

### 本地运行

```bash
# 克隆仓库
git clone <repository-url>

# 进入项目目录
cd findyourseat

# 使用任意 HTTP 服务器运行
python -m http.server 8000
# 或
npx serve .

# 访问 http://localhost:8000
```

## 📖 使用指南

### ✨ 灵活的操作流程

系统支持非线性操作，您可以：
- **自由浏览所有步骤** - 点击步骤指示器随意跳转
- **按需执行操作** - 只在真正需要时才进行验证
- **智能提醒** - 执行操作时自动检查前置条件
- **实时状态** - 步骤完成状态实时更新

### 第一步：上传员工数据

支持的Excel格式示例：

| Number | Role | First Name | Last Name | Email | Job Title |
|--------|------|------------|-----------|--------|-----------|
| 1 | sales | Zhang | Wei | zhang.wei@company.com | Sales Manager |
| 2 | customer advisory | Li | Ming | li.ming@company.com | Customer Advisor |

**支持的文件格式：**
- **.xlsx** - Excel 2007+ 标准格式（推荐）
- **.xls** - Excel 97-2003 兼容格式

**支持的列名识别：**
- 编号：Number, 编号, 序号, 工号
- 角色：Role, 角色, 职位, 岗位
- 姓名：Name, 姓名, First Name, Last Name
- 邮箱：Email, 邮箱, 电子邮件

**文件要求：**
- 文件大小限制：最大10MB
- 推荐使用XLSX标准格式
- 不支持包含宏代码的XLSM文件
- 如有XLSM文件，请转换为XLSX格式后上传

### 第二步：配置会议室

- 设置会议室数量（1-10个）
- 为每个会议室配置桌子数量
- 选择每桌座位数（8-10人）
- 可使用预设模板或导入已保存的配置

### 第三步：智能分配

**灵活的执行方式**：
- 可随时跳转到此步骤查看
- 点击"开始智能分配"才真正执行
- 缺少前置条件时会友好提醒
- 支持重新分配和调整优化

**分配算法流程**：
1. 分析角色分布和人数统计
2. 计算理想的角色分配比例
3. 随机打乱并初步分配
4. 优化调整以提高角色分散度

### 第四步：预览调整

- 查看可视化座位图
- 通过拖拽手动调整座位
- 实时查看角色分布统计
- 支持重新分配和撤销修改

### 第五步：导出结果

支持多种导出格式：
- **按桌分组**：Excel格式，每桌单独分组显示
- **座位清单**：完整的人员座位对照表
- **统计汇总**：角色分布和利用率统计
- **打印版本**：适合打印的HTML格式

## 🎯 核心算法

### 角色分配策略

1. **均匀分布原则**：确保每桌都有不同角色的代表
2. **避免聚集**：限制同一桌子的同角色人员数量
3. **随机性保证**：相同配置下可产生不同的分配结果
4. **优化调整**：通过交换算法持续优化分散度

### 多样性计算

使用Shannon多样性指数评估桌子的角色分布：

```
Diversity = -Σ(pi * log2(pi))
```

其中 pi 是角色 i 在该桌的占比。

## 🛠️ 技术架构

### 前端技术栈

- **HTML5 + CSS3 + JavaScript (ES6+)**
- **Bootstrap 5** - UI框架和响应式设计
- **SheetJS** - Excel文件解析和生成
- **SortableJS** - 拖拽功能实现

### 项目结构

```
findyourseat/
├── index.html           # 主页面
├── css/
│   └── style.css        # 样式文件
├── js/
│   ├── app.js          # 主应用逻辑
│   ├── excel-parser.js  # Excel解析模块
│   ├── algorithm.js     # 分配算法模块
│   ├── visualization.js # 可视化预览模块
│   └── export.js       # 导出功能模块
├── libs/
│   ├── xlsx.full.min.js # Excel处理库
│   └── sortable.min.js  # 拖拽功能库
└── README.md           # 说明文档
```

### 核心模块

#### ExcelParser 模块
- 智能列名识别和数据验证
- 支持多种Excel格式
- 角色标准化和数据清理

#### SeatingAlgorithm 模块
- 多阶段分配算法
- 角色分散度优化
- 约束检查和验证

#### SeatingVisualizer 模块
- Canvas/SVG可视化渲染
- 拖拽交互实现
- 实时统计更新

#### ExcelExporter 模块
- 多工作表Excel生成
- 按需导出不同格式
- 打印友好的HTML生成

## 🎨 UI设计特色

### 外企风格设计
- 简洁现代的界面布局
- 一致的色彩搭配和间距
- 清晰的信息层次和引导流程

### 用户体验优化
- **非线性操作流程** - 支持自由跳转浏览
- **智能前置检查** - 执行时才验证必要条件
- **友好的错误提示** - 清楚告知缺失的步骤
- **实时状态更新** - 步骤完成状态可视化
- **操作确认机制** - 重要操作需用户主动触发

## 🔧 配置管理

### 本地存储
- 使用 localStorage 保存用户配置
- 自动保存最近的会议室设置
- 支持会话恢复

### 配置分享
- JSON格式配置文件导入/导出
- URL参数编码实现一键分享
- 预设模板快速应用

### 数据安全
- 所有数据处理在本地完成
- 不上传任何敏感信息到服务器
- 支持私有部署
- 专注支持标准Excel格式，确保稳定性

## 📱 兼容性

- **现代浏览器支持**：Chrome 70+, Firefox 65+, Safari 12+, Edge 79+
- **移动设备适配**：响应式设计，支持平板和手机访问
- **文件格式支持**：Excel (.xlsx, .xls), CSV

## 🚀 未来规划

- [ ] 增加更多座位布局样式（圆桌、剧院式等）
- [ ] 支持特殊需求人员的座位约束
- [ ] 集成邮件发送功能
- [ ] 添加多语言支持
- [ ] 提供API接口供其他系统集成

## 🤝 贡献指南

欢迎提交 Issue 和 Pull Request 来帮助改进这个项目！

### 开发环境设置

1. Fork 项目
2. 创建功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 打开 Pull Request

### 代码规范

- 使用 ES6+ 语法
- 遵循 JSDoc 注释规范
- 保持代码模块化和可维护性
- 确保响应式设计兼容性

## 📄 许可证

本项目采用 MIT 许可证 - 详见 [LICENSE](LICENSE) 文件

## 🙏 致谢

- [Bootstrap](https://getbootstrap.com/) - UI框架
- [SheetJS](https://sheetjs.com/) - Excel处理
- [SortableJS](https://sortablejs.github.io/Sortable/) - 拖拽功能
- [Bootstrap Icons](https://icons.getbootstrap.com/) - 图标库

---

如有问题或建议，请通过 Issue 或邮件联系我们。祝您使用愉快！ 🎉
