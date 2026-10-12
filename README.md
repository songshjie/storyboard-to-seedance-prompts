# storyboard-to-seedance-prompts

豆包（Doubao）Agent 技能：把用户提供的视频分镜表（xlsx / csv / 飞书表格 / 分镜文档）转化为可直接用于 **Seedance / 即梦** 等 AI 视频生成的详细导演版提示词，并交付为**飞书在线文档链接**。

分镜表内容单薄时，会结合全片剧情自动补齐镜头语言、表演细节、光线与声音设计（不改剧情、不加角色、不改台词原文）。

## 目录结构

```
storyboard-to-seedance-prompts/
├── SKILL.md                              # 技能主文件：完整工作流
├── scripts/
│   └── parse_storyboard.py               # 分镜表解析脚本（自动识别列、生成时间码、抽取人物道具）
└── references/
    ├── format-template.md                # 飞书文档结构模板 + 逐镜行格式
    └── director-prompt-guide.md          # 画面提示词七要素、运镜语言、声音设计、丰富策略
```

## 使用方式

1. 将技能文件夹放入豆包的 `workspace/.user_skills/` 目录；
2. 直接向豆包发送分镜表（xlsx / csv / 飞书表格链接 / 分镜文档），并说明需要 Seedance 提示词；
3. 豆包自动产出逐镜导演版提示词，交付飞书文档链接，可直接逐镜复制到 Seedance / 即梦生成。

## 依赖

- Python 3 + openpyxl（解析 xlsx）
- 豆包环境的 lark-doc / sheet 技能（创建飞书文档与读取飞书表格）
