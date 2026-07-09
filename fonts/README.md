# 中文字体

本目录的中文字体供 ONLYOFFICE Document Server 把 `.docx` 转 PDF 时渲染中文使用。
通过 `docker-compose.yml` 只读挂载到容器 `/usr/share/fonts/truetype/custom/`；
`start.sh` 在首次启动或容器重建后会自动 `fc-cache` + `documentserver-generate-allfonts.sh` 完成注册。

| 文件 | 字体 |
|------|------|
| simsun.ttc | 宋体 / 新宋体 |
| simfang.ttf | 仿宋 |
| simhei.ttf | 黑体 |
| simkai.ttf | 楷体 |
| msyh.ttc | 微软雅黑 |

> 字体来源为 Windows 系统字体，版权归 Microsoft / 各字体厂商所有。
> 仅用于本项目文档转换的内部渲染，请勿从本仓库另行分发。
