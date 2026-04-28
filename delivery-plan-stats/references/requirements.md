# 依赖项说明

## Python 依赖

本技能需要以下 Python 包：

```
pandas==2.0.0
openpyxl==3.1.2
matplotlib==3.7.0
numpy==1.24.0
```

## 安装方式

```bash
pip install pandas openpyxl matplotlib numpy
```

## 中文字体支持

生成图片时需要中文字体支持，脚本会自动尝试以下字体：
- Noto Sans CJK SC
- SimHei
- Microsoft YaHei
- Arial Unicode MS

如需安装中文字体（Linux环境）：
```bash
# Ubuntu/Debian
sudo apt-get install fonts-noto-cjk

# CentOS/RHEL
sudo yum install google-noto-sans-cjk-fonts
```

## 验证安装

```python
import pandas as pd
import openpyxl
import matplotlib.pyplot as plt
import numpy as np

print("所有依赖已正确安装")
```
