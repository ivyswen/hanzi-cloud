"""
stopwords.py — 停用词管理模块
负责停用词的加载（从外部文件）与持久化保存。
"""

import os
from pathlib import Path

# 内置默认停用词，首次运行时写入文件作为初始内容
_DEFAULT_STOPWORDS: set[str] = {
    "的", "了", "在", "是", "我", "有", "和", "就", "不", "人",
    "都", "一", "一个", "上", "也", "很", "到", "说", "要", "去",
    "你", "会", "着", "没有", "看", "好", "自己", "这", "他", "她",
    "它", "们", "那", "什么", "我们", "他们", "她们", "但", "还",
    "与", "对", "中", "为", "以", "及", "等", "将", "把", "被",
}

# 停用词文件路径（与程序同目录）
_STOPWORDS_FILE = Path(__file__).parent / "stopwords.txt"


def load_stopwords() -> set[str]:
    """
    从 stopwords.txt 加载停用词集合。
    若文件不存在，则用内置默认词集创建文件后返回。
    """
    if not _STOPWORDS_FILE.exists():
        # 首次运行：写入默认词集供用户参考和编辑
        save_stopwords(_DEFAULT_STOPWORDS)
        return set(_DEFAULT_STOPWORDS)

    words: set[str] = set()
    with open(_STOPWORDS_FILE, encoding="utf-8") as f:
        for line in f:
            word = line.strip()
            if word:  # 跳过空行
                words.add(word)

    # 若文件为空，回退到默认值
    return words if words else set(_DEFAULT_STOPWORDS)


def save_stopwords(words: set[str]) -> None:
    """
    将停用词集合写回 stopwords.txt（每行一词，按字典序排序）。
    """
    with open(_STOPWORDS_FILE, "w", encoding="utf-8") as f:
        for word in sorted(words):
            f.write(word + "\n")
