# Excel 角色卡

<p align="center"><img src="icon.png" width="96" height="96" alt="Excel 角色卡"></p>

> 把一张 Excel 角色表变成 DSH 里的常驻角色：上传后，助手会按表里的性别、年龄、身高、性格、身份、说话风格等设定，以第一人称持续扮演。

仓库：[github.com/wuliao117/dsh-excel-character-card](https://github.com/wuliao117/dsh-excel-character-card)

[![test](https://github.com/wuliao117/dsh-excel-character-card/actions/workflows/test.yml/badge.svg)](https://github.com/wuliao117/dsh-excel-character-card/actions/workflows/test.yml)

## 为什么需要它

想让 AI 扮演一个特定角色时，通常要先把人设贴进对话，角色一多就难以管理，换会话还会丢失。这个插件把角色档案放在 Excel 里：一张表可以写一个角色，也可以写一整组角色；导入后角色卡保存在本机，按会话启用，之后每一轮回复都会自动带上角色设定。

和“每次把设定粘进提示词”相比：

- 角色表可以反复编辑、批量维护；
- 角色卡只导入一次，换会话仍可复用；
- 启用/停用是显式操作，不会污染普通对话；
- 系统提示词会持续注入角色档案，不依赖用户每轮提醒。

## 功能

- **两种 Excel 布局**：键值表（A 列字段、B 列内容）和角色表（第一行表头、每行一个角色）。
- **多角色**：一个工作簿可以包含多个工作表、多个角色；空白行分隔的多个键值块也会被识别。
- **中文/英文字段名**：姓名、性别、年龄、身高、体重、生日、种族、身份、外貌、性格、喜好、讨厌、恐惧、背景、动机、技能、弱点、说话风格、口头禅、关系、秘密、规则等，中英文别名都支持。
- **未知字段不丢**：表里多出来的列/键会进入“其他设定”，一起注入提示词。
- **按会话启用**：每个会话可以扮演不同角色；也可以设置全局默认角色。
- **可视化面板**：输入框右侧的「角色卡」按钮可以拖入 `.xlsx` / `.xlsm`、预览、启用、停用、删除。
- **给模型用的工具**：`excel_character_card`，支持 `list` / `import` / `preview` / `get` / `activate` / `deactivate` / `delete`。

## Excel 怎么写

### 方式一：键值表（推荐，适合单角色）

| A 列 | B 列 |
| --- | --- |
| 姓名 | 林晚 |
| 性别 | 女 |
| 年龄 | 24 |
| 身高 | 168 cm |
| 身份 | 独立咖啡师 |
| 外貌 | 黑色长发，琥珀色眼睛…… |
| 性格 | 表面温和疏离，熟了之后话多…… |
| 说话风格 | 语速偏慢，句子短…… |
| 背景 | 曾在香水公司做调香师…… |
| 规则 | 不要主动说出自己的秘密…… |

A 列是字段名，B 列是内容。字段顺序不限，行数不限；用空白行分隔可以写多个角色（每个角色都要有“姓名”）。

### 方式二：角色表（适合多角色）

| 姓名 | 性别 | 年龄 | 身高 | 性格 | 身份 | 说话风格 | 背景 |
| --- | --- | --- | --- | --- | --- | --- | --- |
| 沈砚 | 男 | 29 | 185 cm | 冷静克制，毒舌但靠谱 | 刑警 | 低音、短句 | 因为一桩旧案从法医转做刑警 |
| 阿澄 | 女 | 26 | 160 cm | 热情开朗，擅长察言观色 | 花店老板 | 语速快，爱用感叹号 | 在老街长大 |
| 顾言 | 男 | 34 | 178 cm | 温和有礼，边界感强 | 心理医生 | 措辞精确、声音轻 | 曾因一次误诊离开公立医院 |

第一行是表头，之后每行一个角色。表头里不认识的列会作为“其他设定”保留。

### 其他写法

- 横排键值：`姓名 | 林晚 | 性别 | 女 | 年龄 | 24 | 身高 | 168cm`。
- 单元格内写 `字段：内容`。
- 合并单元格：值列纵向合并时会自动填充到每一行。
- 日期：常见 Excel 日期格式会转成 `YYYY-MM-DD`。

仓库内的 [`sample/角色卡示例.xlsx`](sample/角色卡示例.xlsx) 同时包含上面两种布局，可以直接拖进面板试。

## 安装

### 从 GitHub 安装（推荐）

```sh
dsh plugin --profile web add github:wuliao117/dsh-excel-character-card
dsh --profile web
```

装完重启 web（bundle 层栈在启动时合成）。

### 桌面端（Electron）

桌面端 profile 由应用管理，推荐让 DSH 自己安装 GitHub bundle：

> 用 `plugin_manager` 从 `github:wuliao117/dsh-excel-character-card` 安装 bundle，装完告诉我是否需要重启。

安装后如果面板没有出现，完全退出应用（含系统托盘）再重新打开。

### 本地开发安装

```sh
dsh plugin --profile web add D:\gz\excel-character-card
```

本地目录安装要求目录里有 `package.json#dsh.bundle` 且 `lib/` 产物已在库；本插件没有第三方运行时依赖，因此不需要额外 link SDK。装完重启 web。

## 使用

1. 在输入框右侧点「角色卡」。
2. 把 `.xlsx` / `.xlsm` 拖进面板，或点击选择文件。
3. 在角色列表里点「启用」。面板顶部会显示“当前会话角色”。
4. 直接开始聊天；角色设定会随每一轮请求注入。
5. 想恢复普通助手：点「停用」。

也可以直接对模型说：

> 读取 `D:\...\角色卡.xlsx`，导入后扮演林晚。

模型会调用 `excel_character_card`：`action=import`、`path=...`、`activate=true`。

## 数据与隐私

- 角色卡保存在 `%USERPROFILE%\.dsh\excel-character-card\state.json`（或 `$DSH_HOME` 对应目录）。
- 上传的文件只在 Host 内存中解析，不会写入其他位置。
- 角色档案会进入发给模型的系统提示词；不要在 Excel 里放 API key、密码等敏感信息。
- 插件只监听本机 DSH 的 HTTP 服务，不向外部发送数据。

## 工具

| 工具 | 说明 |
| --- | --- |
| `excel_character_card` | `list` 列出角色；`import` 导入（`path`，可选 `sheet`、`activate`）；`preview` 只解析不保存；`get` 查看档案与提示词；`activate` / `deactivate` 控制当前会话；`delete` 删除。 |

## 已知限制

- 仅支持 `.xlsx` / `.xlsm`；旧版二进制 `.xls` 请先另存为 `.xlsx`。
- 单个文件上限 20 MB，最多保存 200 张角色卡。
- 角色扮演的稳定性取决于模型是否遵守系统提示词；角色卡写得越具体，效果越稳定。
- 面板按会话记录启用状态；删除角色卡会同时清理指向它的启用记录。
- 角色卡注入会占用 token；过长的背景/关系建议精简到关键信息。

## 开发与测试

```sh
node test/parse.test.mjs
node test/host.test.mjs
```

- `lib/xlsx.mjs`：零依赖 XLSX 读取（ZIP + XML + 共享字符串 + 合并单元格 + 日期格式）。
- `lib/character.mjs`：字段别名、布局识别、角色卡提取、提示词生成。
- `lib/index.js`：Host 半（工具、系统提示词、HTTP API、持久化）。
- `lib/client.js`：浏览器半（输入框按钮 + 角色卡面板）。

## 插件管理

已装插件可用 DSH 自带的插件管理器（设置 → 插件）启停和查看；本插件是标准 `dsh.bundle` 包，卸载时移除该 bundle 即可。

## License

MIT。XLSX 解压使用 vendored [fflate](https://github.com/101arrowz/fflate)（MIT）。
