# 栖账 · 本地离线记账

根据 V1.0 PRD 开发的移动端 MVP。基础交互参考 BeeCount，采用独立的绿色视觉。应用不需要账号，核心财务数据保存在设备本地；界面为移动端优先设计，可作为 PWA 安装，也已准备 Capacitor Android 封装配置。

## 已完成

- 明细：首页月度收支、日期分组账目与日历入口；提醒统一进入右上角消息中心
- 账目：按日分组、月度切换、关键词搜索、详情、编辑和永久删除
- 记账：支出、收入、转账、二级分类与四则计算器
- 账户：资产/负债账户、本地余额实时计算、新增、编辑和安全删除
- 账本与分类：账本完整管理、一级/二级分类新增编辑、引用保护与拖动排序
- 洞察：日/月/年汇总、周期趋势、支出与收入分类占比，并可下钻查看二级分类与相关账目
- 分类图标：内置 48 个常用图标，并支持自定义输入
- 资产：净资产环图、资产/负债分组、账户流水，以及活钱管理、保险保障、长期投资、稳健理财标签配置与占比可视化
- 预算：各账本独立设置分类月预算，首页查看额度、使用情况与剩余额度
- Excel：导出包含账目、账户、账本、分类、预算等工作表的 `.xlsx`；导入前预览并合并到本地
- 排序同步：“我的”中的账本、账户及分类顺序自动同步到前台切换列表、资产页和记账页
- 多账本切换、新建账本、日历明细、消息中心、金额隐藏、深色模式
- LocalStorage 本地持久化、Service Worker 离线缓存、PWA 安装清单

当前版本不包含 AI/OCR/语音记账、云同步、自建服务、周期记账和借贷管理。

版本历史及 APK 校验值见 [RELEASES.md](RELEASES.md)。

## 设计参考

基础信息架构参考 [TNT-Likely/BeeCount](https://github.com/TNT-Likely/BeeCount)。BeeCount 使用自定义软件许可：个人使用和学习研究免费，商业使用需要获得作者授权。参考源码仅用于本地设计参考，不会上传到本仓库或打包进应用。

## 本地运行

```powershell
node scripts/serve.mjs
```

浏览器打开 `http://127.0.0.1:4173`。首次打开为干净账本；在“我的 → 本地数据”中可载入体验数据，或直接访问 `http://127.0.0.1:4173/?demo=1`。

## Android 封装

电脑安装 JDK 17 和 Android Studio/Android SDK 后执行：

```powershell
npm install
npm run android:init
npm run android:sync
npm run android:open
```

然后在 Android Studio 中运行或生成 APK。Android 包名为 `com.qizhang.offline`。

`npm run build` 会把离线应用资源生成到 `www/`，Android 封装只读取该目录，不会混入项目源码。

## 构建说明

仓库忽略 `node_modules/`、工具链、生成的 `www/`、Android 构建缓存和本机 SDK 路径。克隆后运行 `npm install`、`npm run android:sync`，再使用 Android Studio 或 Gradle 构建 APK。

已归档的 `dist/Qizhang-v1.0.apk` 至 `dist/Qizhang-v1.9.apk` 和对应 SHA-256 文件保留为版本产物；早期 APK 没有对应源码快照，逐版说明以 [RELEASES.md](RELEASES.md) 标注的可考证信息为准。

## 数据说明

当前 MVP 使用浏览器或 WebView 的本地存储。清空应用数据、浏览器站点数据或卸载应用会永久删除账务数据。PRD 中的完整备份恢复原本属于后续版本，因此本版不提供云同步或服务器备份。
