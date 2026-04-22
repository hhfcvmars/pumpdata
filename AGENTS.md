# AGENTS.md

## 项目概述

pump-data — 一个用于处理和编辑胰岛素泵历史数据的 React 单页应用。支持导入 JSON/Excel 文件，在表格中查看、编辑、筛选泵事件记录，并导出修改后的数据。

## 技术栈

- React 19 + Vite 7（纯 JavaScript，无 TypeScript）
- 构建工具：Vite（`vite.config.js`）
- 包管理：npm
- 代码规范：ESLint 9（flat config，`eslint.config.js`）
- 依赖库：xlsx（Excel 读写）、file-saver（文件下载）

## 项目结构

```
├── src/
│   ├── App.jsx          # 主应用组件（所有业务逻辑集中于此，~2100 行）
│   ├── App.css          # 样式
│   ├── main.jsx         # 入口文件
│   └── index.css        # 全局样式
├── public/              # 静态资源
├── index.html           # HTML 模板
├── vite.config.js       # Vite 配置
├── eslint.config.js     # ESLint 配置
└── package.json
```

## 常用命令

```bash
npm run dev       # 启动开发服务器
npm run build     # 生产构建
npm run preview   # 预览生产构建
npm run lint      # ESLint 检查
```

## 关键业务概念

- 设备事件（Device Event）：用十六进制编码表示泵的各类事件（报警、状态变更等），通过 eventPort/eventType/eventLevel 三段组合
- 大剂量（Bolus）：胰岛素注射的大剂量计算，基于时间差和速率
- 基础率（Basal）：持续输注的基础胰岛素速率，输注量 = 时间差 × (basal / 160) U/hr
- 数据流：导入文件 → 解析 → 计算衍生字段（bolusSum、presetBasalUnitPerHour）→ 表格编辑 → 导出

## 编码规范

- 使用 JSX（`.jsx` 扩展名），非 TypeScript
- ESLint 规则：未使用变量报错（大写开头或下划线开头的变量名除外）
- React Hooks 规范 + React Refresh 规范
- 组件和工具函数目前全部在 `App.jsx` 中，无独立模块拆分
- 中文注释和中文 UI 标签
