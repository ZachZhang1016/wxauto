"""
交互式选择要监听的微信群/联系人，并生成 config.json。

用法：
    python select_chats.py

运行前请确保：
1. 已 `pip install -e .`（在仓库根目录）
2. 微信桌面客户端已登录
"""
import json
import os
import re

from wxauto import WeChat

from monitor import DEFAULT_SUMMARY_SCHEDULE, DEFAULT_WEEKLY_SCHEDULE

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.json")

# 重新跑配置向导时原样保留的高级配置项（向导不问这些，只能手动编辑 config.json）
PRESERVED_ADVANCED_KEYS = (
    "summary_model",
    "notify_to",
    "push_summary",
    "signal_alert",
)


def split_list(raw):
    """按中英文逗号、顿号分隔输入，兼容中文输入法打出的全角逗号"""
    return [s.strip() for s in re.split(r"[,，、]", raw) if s.strip()]


def ask_senders(wx, chat_name):
    members = None
    try:
        wx.ChatWith(chat_name)
        members = wx.GetGroupMembers()
    except Exception:
        members = None

    if members:
        print(f'\n  "{chat_name}" 的群成员：')
        for i, m in enumerate(members, 1):
            print(f"    {i}. {m}")
        raw = input(
            "  只记录哪些人的发言？（直接回车 = 全部；否则输入上面的编号，多个用逗号分隔）: "
        ).strip()
        if not raw:
            return []
        tokens = split_list(raw)
        senders = []
        for t in tokens:
            if t.isdigit() and 1 <= int(t) <= len(members):
                senders.append(members[int(t) - 1])
            else:
                senders.append(t)  # 支持直接输入名字（不在编号列表里也行）
        return senders

    # 不是群聊，或没能取到成员列表（比如是1对1好友聊天），退回手动输入名字
    raw = input(
        f'  "{chat_name}" 是否只记录特定成员的发言？'
        f"（直接回车 = 记录全部成员；否则输入成员昵称，多个用逗号分隔）: "
    ).strip()
    return split_list(raw)


def ask_language():
    raw = input(
        "你的微信客户端界面语言：1) 简体中文  2) 繁体中文  3) English  [默认1]: "
    ).strip()
    return {"2": "cn_t", "3": "en"}.get(raw, "cn")


def main():
    language = ask_language()
    print("正在连接微信客户端...")
    wx = WeChat(language=language)

    sessions = wx.GetSessionList(reset=True)
    names = list(sessions.keys())

    print("\n当前会话列表中的群/联系人：")
    for i, name in enumerate(names, 1):
        unread = sessions[name]
        tag = f"（{unread}条未读）" if unread else ""
        print(f"  {i}. {name}{tag}")
    print(
        "\n如果要监听的群/联系人不在上面列表中（比如很久没聊天了），"
        "可以直接输入完整的名称。"
    )

    raw = input(
        "\n请输入要监听的编号和/或名称，用逗号分隔（例如: 1,3,张三）: "
    ).strip()
    tokens = split_list(raw)

    chosen = []
    for t in tokens:
        if t.isdigit() and 1 <= int(t) <= len(names):
            chosen.append(names[int(t) - 1])
        else:
            chosen.append(t)
    # 去重，保持顺序
    seen = set()
    chosen = [c for c in chosen if not (c in seen or seen.add(c))]

    if not chosen:
        print("未选择任何聊天对象，退出。")
        return

    chats = []
    for name in chosen:
        senders = ask_senders(wx, name)
        chats.append({"name": name, "senders": senders})

    poll_raw = input(
        "\n每隔多少分钟记录一次消息？（建议5分钟左右，太长可能会漏消息）[默认5]: "
    ).strip()
    poll_interval_minutes = int(poll_raw) if poll_raw.isdigit() else 5

    media_raw = input(
        "是否自动保存图片/文件/语音转文字？(y/N，默认N，只影响记录中的媒体消息): "
    ).strip().lower()
    save_media = media_raw == "y"

    # 摘要时间表、交易信号提醒等是相对固定的高级配置，已经配置过就不要覆盖，
    # 只在第一次生成配置时写入默认值
    existing_config = {}
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            existing_config = json.load(f)

    config = {
        "language": language,
        "poll_interval_minutes": poll_interval_minutes,
        "chats": chats,
        "summary_schedule": existing_config.get("summary_schedule")
        or DEFAULT_SUMMARY_SCHEDULE,
        "weekly_summary_schedule": existing_config.get("weekly_summary_schedule")
        or DEFAULT_WEEKLY_SCHEDULE,
        "save_media": {
            "savepic": save_media,
            "savefile": save_media,
            "savevoice": save_media,
        },
    }
    for key in PRESERVED_ADVANCED_KEYS:
        if key in existing_config:
            config[key] = existing_config[key]

    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)

    print(f"\n配置已保存到 {CONFIG_PATH}")
    print(f"监听对象: {[c['name'] for c in chats]}")
    print(f"记录频率: 每 {poll_interval_minutes} 分钟")
    print("摘要时间表: " + ", ".join(f"{e['label']}@{e['time']}" for e in config["summary_schedule"]))
    ws = config["weekly_summary_schedule"]
    print(f"周报时间表: 每周{ws['day']} {ws['time']}")
    print("（这些时间表是高级配置，如需调整可以直接编辑 config.json）")
    print("\n接下来可以运行:")
    print("  python monitor.py --once     # 手动跑一次")
    print("  python monitor.py            # 常驻运行")


if __name__ == "__main__":
    main()
