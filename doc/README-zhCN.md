# zotero-thief

[![Zotero 7](https://img.shields.io/badge/Zotero-7-green?style=flat-square&logo=zotero&logoColor=CC2936)](https://www.zotero.org)
[![TypeScript](https://img.shields.io/badge/TypeScript-supported-blue?style=flat-square&logo=typescript)](https://www.typescriptlang.org)

![效果预览](../effect.png)

zotero-thief 是一个 Zotero 插件，用来把阅读器里选中的文本快速切换成一个轻量的小说阅读界面。

## 亮点

- 在阅读器里选中文本后按 r，启动小说阅读界面。
- 使用 w / q / e 前进、后退，以及隐藏或恢复替换内容。
- 在设置页管理书架、选中的小说、语言和快捷键。

## 安装

1. 安装 Zotero 7 和 Node.js LTS。
2. 克隆本仓库。
3. 运行 `npm install`。

## 开发

运行 `npm start` 启动开发流程，支持热重载。

运行 `npm run build` 生成生产包。

## 配置

打开插件设置页后，可以配置：

- 插件语言
- 小说字体大小
- 图书存储目录
- 当前选中的小说
- 快捷键

## 项目结构

- [src/modules/novel.ts](../src/modules/novel.ts) 包含主要的小说替换与快捷键逻辑。
- [src/modules/preferenceScript.ts](../src/modules/preferenceScript.ts) 管理设置页界面。
- [src/modules/examples.ts](../src/modules/examples.ts) 负责阅读器事件与上下文动作连接。

## 许可证

AGPL-3.0-or-later。