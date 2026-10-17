# ecw2pptx
> 一键将畅言智慧课堂 `.ecw` 课件文件转换为标准 `.pptx` 演示文稿

## 简介
`.ecw` 是**畅言智慧课堂**使用的私有课件封装格式，并非标准PPT。本工具直接解析 ecw 内部资源与私有XML，重建标准OOXML，导出可被 PowerPoint、WPS、LibreOffice 正常打开的 `.pptx` 文件。

畅言智慧课堂官网：https://www.changyan.com/

## 用法
```bash
python3 ecw2pptx.py <input.ecw> [output.pptx]
```

## 功能
- 解析 ecw 包内 slide / layout / master / theme 的私有 XML
- 将形状(文本/图片/组合/自定义几何/渐变/阴影)映射为 OOXML DrawingML
- 重建 `[Content_Types].xml` / `_rels` 关系链 / `presentation.xml`
- 内嵌 resource 图片，输出标准 OOXML `.pptx`(可由 PowerPoint / WPS / LibreOffice 打开)

## 依赖
仅依赖 Python 标准库：
- `xml.etree.ElementTree`
- `zipfile`

> 无需安装第三方pip包，Python3环境可直接运行

## 注意事项
1. `.ecw` 为安徽科大讯飞旗下畅言智慧课堂私有课件格式，本项目仅用于**个人学习、格式研究**。
2. 请勿用于商业用途、批量解密或侵权分发课件文件，请遵守版权相关法律法规。
3. 部分特殊动画、课堂交互组件、特殊音视频控件无法完全转换，仅转换静态页面内容。

