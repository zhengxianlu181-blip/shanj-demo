# 闪记 PC 功能演示

根目录的 `index.html` 是思必驰闪记 `release-3.3.0-pc` 真实 Flutter Web 构建，可直接通过 GitHub Pages 分享：

<https://zhengxianlu181-blip.github.io/shanj-demo/>

当前构建包含：

- 首页“开始录音”和“文件导入”入口
- 实时录音、暂停/继续、结束确认、计时和逐句转写
- 说话人、时间戳、字幕设置和字幕浮窗入口
- MP3、WAV、M4A、MP4、MOV 文件选择/拖拽、播放预览、上传、转写、保存和失败重试
- 实时录音和文件导入共用的录音详情页

## 设计评审版

如果需要查看录音工作流的交互状态和“自定义总结”弹窗，可打开：[设计评审版](./design-review.html)。该页面用于 UI 评审，不替代根目录的真实 Flutter 功能构建。

在线页面受浏览器能力限制：真实麦克风、登录、ASR 服务和上传接口仍需要对应环境与权限；Windows 原生字幕浮窗的置顶、拖动、最小化和完整设置请在源码仓库运行 PC 端：

```text
cd apps/app_main
flutter run -d windows
```

根目录的真实 Flutter 构建用于功能演示；`design-review.html` 用于页面布局、状态和交互评审。
