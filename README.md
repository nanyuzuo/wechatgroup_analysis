# 马哥教育大模型1期微信群成员分析工具 v1.0

这是一个功能强大的“马哥教育大模型1期”微信群成员分析工具，可以自动统计和可视化群成员的地理分布情况，生成美观的数据报告。该工具采用现代化UI设计，支持多种数据可视化方式，让群成员分析变得简单而专业。

## 🌟 功能特点

- 自动获取微信群成员信息，支持多种获取方式
- 智能解析成员地理位置信息（支持省份、城市、直辖市等）
- 自动识别并统计马哥教育管理员
- 生成美观的统计报告，包含：
  - 群成员总体分布分析
  - 地理位置热力图
  - 省份分布统计
  - 城市分布统计
  - 管理员统计
- 支持多种可视化图表：
  - 地理分布地图
  - 成员分布柱状图
  - 数据统计卡片
  - 现代化UI设计

## 🔧 环境要求

- Windows 操作系统
- Python 3.x
- PC版微信（推荐使用 3.9.11.17 版本）
- Windows默认黑体字体（用于生成图片）

## 📦 安装依赖

```bash
pip install -r requirements.txt
```

主要依赖包括：
- wxauto==3.9.11.17.5（微信自动化）
- pandas==2.1.1（数据处理）
- matplotlib==3.8.0（图表生成）
- pillow==10.0.1（图像处理）
- geopandas（地理数据处理）
- wordcloud（词云生成）
- 其他辅助包：numpy, pyperclip, requests等

## 🚀 使用方法

1. 确保PC版微信已登录并保持运行状态
2. 打开要分析的微信群聊窗口
3. 运行程序：
   ```bash
   python wechat_group_analysis.py
   ```
4. 根据提示输入要分析的微信群名称
5. 等待程序自动完成分析和报告生成

## 📊 输出结果

程序会自动生成美观的分析报告，包含：
- 群成员统计信息
- 地理分布热力图
- 省份和城市分布统计
- 管理员信息统计
- 其他可视化图表

所有结果都会以图片形式保存，便于分享和查看。

## ⚠️ 注意事项

1. 群成员命名规范：
   - 普通成员格式：`学号-城市-昵称`
   - 管理员昵称中需包含"马哥"字样

2. 运行环境注意事项：
   - 确保微信窗口处于正常状态（未最小化）
   - 运行时不要操作微信窗口
   - 确保系统中安装了所需字体
   - 保持网络连接（用于获取地图数据）

3. 数据获取说明：
   - 程序支持多种方式获取群成员信息
   - 如果主要方式失败，会自动尝试备选方案
   - 支持对直辖市、省份、城市的智能识别

## 🔍 故障排除

如果遇到问题，请检查：

1. 微信连接问题：
   - 确认微信版本是否为推荐版本
   - 检查微信是否正常登录
   - 确保群聊窗口正常打开

2. 数据获取问题：
   - 验证群名称是否正确
   - 确认您是否为群成员
   - 检查群聊是否有最近的聊天记录

3. 图表生成问题：
   - 检查依赖包是否正确安装
   - 确认系统字体是否完整
   - 验证磁盘空间是否充足

## 🤝 技术支持

如果需要帮助，请：
1. 检查以上故障排除步骤
2. 确认所有依赖包已正确安装
3. 验证运行环境是否满足要求

## 📝 版本说明

当前版本：v1.0
- 完整的群成员数据分析功能
- 现代化UI设计
- 智能地理位置识别
- 多种可视化展示方式
- 稳定的数据获取机制 