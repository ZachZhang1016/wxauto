"""
用弹窗界面选择要监听的微信群/联系人，并生成 config.json。
不用在终端里手打群名/人名——群和人都从真实列表里点选。

用法：
    python select_chats_gui.py
"""
import json
import os
import re
import tkinter as tk
from tkinter import messagebox, ttk

from wxauto import WeChat

from monitor import DEFAULT_SUMMARY_SCHEDULE, DEFAULT_WEEKLY_SCHEDULE
from select_chats import PRESERVED_ADVANCED_KEYS

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.json")


def split_names(raw):
    """按换行/中英文逗号/顿号切分人名，兼容一行输入多个名字的习惯写法"""
    parts = re.split(r"[\n,，、]", raw)
    return [p.strip() for p in parts if p.strip()]


class LanguageDialog(tk.Toplevel):
    """启动时先问微信客户端界面语言"""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("微信界面语言")
        self.resizable(False, False)
        self.result = "cn"

        tk.Label(self, text="你的微信客户端界面语言是？", padx=16, pady=12).pack()
        self.var = tk.StringVar(value="cn")
        for label, value in [("简体中文", "cn"), ("繁体中文", "cn_t"), ("English", "en")]:
            tk.Radiobutton(self, text=label, variable=self.var, value=value).pack(
                anchor="w", padx=24
            )
        tk.Button(self, text="连接微信", command=self._ok, width=14).pack(pady=12)

        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.grab_set()

    def _ok(self):
        self.result = self.var.get()
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()


class MemberPickerDialog(tk.Toplevel):
    """
    选择某个群/联系人要记录哪些人的发言。
    能拿到真实群成员列表就直接点选；拿不到（比如是1对1好友，或者取不到）就退回
    一个文本框，一行填一个人的昵称——不用逗号分隔，emoji、括号等特殊符号都不影响。
    """

    def __init__(self, parent, chat_name, members):
        super().__init__(parent)
        self.title(f'选人 - "{chat_name}"')
        self.geometry("380x420")
        self.result = []  # [] 表示记录全部人
        self._members = members

        tk.Label(
            self,
            text=f'"{chat_name}"\n只记录哪些人的发言？不选/留空 = 记录全部人',
            padx=12,
            pady=10,
            justify="left",
        ).pack(fill="x")

        if members:
            frame = tk.Frame(self)
            frame.pack(fill="both", expand=True, padx=12)
            scrollbar = tk.Scrollbar(frame)
            scrollbar.pack(side="right", fill="y")
            self.listbox = tk.Listbox(
                frame, selectmode=tk.EXTENDED, yscrollcommand=scrollbar.set
            )
            for m in members:
                self.listbox.insert(tk.END, m)
            self.listbox.pack(side="left", fill="both", expand=True)
            scrollbar.config(command=self.listbox.yview)
            tk.Label(
                self, text="（可以按住 Ctrl 或 Shift 多选）", fg="gray"
            ).pack(anchor="w", padx=12)
            self.text = None
        else:
            tk.Label(
                self,
                text="（没能取到群成员列表，改成手动输入：一行一个昵称，或用逗号/顿号分隔也行，留空=全部）",
                fg="gray",
                wraplength=340,
                justify="left",
            ).pack(anchor="w", padx=12)
            self.text = tk.Text(self, height=14)
            self.text.pack(fill="both", expand=True, padx=12, pady=(0, 8))
            self.listbox = None

        btns = tk.Frame(self)
        btns.pack(pady=10)
        tk.Button(btns, text="确定", width=12, command=self._ok).pack(
            side="left", padx=6
        )

        self.protocol("WM_DELETE_WINDOW", self._ok)
        self.grab_set()
        self.wait_window(self)

    def _ok(self):
        if self.listbox is not None:
            self.result = [self._members[i] for i in self.listbox.curselection()]
        elif self.text is not None:
            raw = self.text.get("1.0", tk.END)
            self.result = split_names(raw)
        self.destroy()


def try_get_group_members(wx, chat_name):
    try:
        wx.ChatWith(chat_name)
        return wx.GetGroupMembers() or None
    except Exception:
        return None


class MainApp(tk.Tk):
    def __init__(self, wx, sessions):
        super().__init__()
        self.wx = wx
        self.session_names = list(sessions.keys())
        self.config_result = None

        self.title("wxauto 监听配置")
        self.geometry("460x520")

        tk.Label(
            self,
            text="选择要监听的群/联系人（可按住 Ctrl/Shift 多选）：",
            padx=12,
            pady=8,
        ).pack(anchor="w")

        frame = tk.Frame(self)
        frame.pack(fill="both", expand=True, padx=12)
        scrollbar = tk.Scrollbar(frame)
        scrollbar.pack(side="right", fill="y")
        self.listbox = tk.Listbox(
            frame, selectmode=tk.EXTENDED, yscrollcommand=scrollbar.set
        )
        for name, unread in sessions.items():
            tag = f"  ({unread}条未读)" if unread else ""
            self.listbox.insert(tk.END, f"{name}{tag}")
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.listbox.yview)

        manual_frame = tk.Frame(self)
        manual_frame.pack(fill="x", padx=12, pady=(6, 0))
        tk.Label(manual_frame, text="列表里没有的，手动加一个名称：").pack(
            side="left"
        )
        self.manual_entry = tk.Entry(manual_frame)
        self.manual_entry.pack(side="left", fill="x", expand=True, padx=6)
        tk.Button(manual_frame, text="添加", command=self._add_manual).pack(
            side="left"
        )
        self.manual_names = []
        self.manual_label = tk.Label(self, text="", fg="gray", anchor="w")
        self.manual_label.pack(fill="x", padx=12)

        options = tk.LabelFrame(self, text="记录设置", padx=10, pady=8)
        options.pack(fill="x", padx=12, pady=10)
        tk.Label(options, text="每隔几分钟记录一次消息:").pack(side="left")
        self.interval_var = tk.IntVar(value=5)
        tk.Spinbox(
            options, from_=1, to=60, width=4, textvariable=self.interval_var
        ).pack(side="left", padx=(4, 12))
        self.media_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            options, text="自动保存图片/文件/语音转文字", variable=self.media_var
        ).pack(side="left")
        tk.Label(
            self,
            text="（摘要触发时间表用的是默认的盘前/盘中/盘后/夜盘/每日总结，"
            "如需调整可以直接编辑 config.json 里的 summary_schedule）",
            fg="gray",
            wraplength=430,
            justify="left",
        ).pack(fill="x", padx=12)

        tk.Button(
            self, text="下一步：分别选择每个群要监听的人 →", command=self._next
        ).pack(pady=10)

        self.protocol("WM_DELETE_WINDOW", self._cancel)

    def _add_manual(self):
        name = self.manual_entry.get().strip()
        if not name:
            return

        if name not in self.session_names:
            # 输入的名字不完全匹配任何真实会话名（比如漏打了emoji/后缀），
            # 尝试模糊匹配，避免存进一个微信里根本找不到的名字
            candidates = [s for s in self.session_names if name in s]
            if len(candidates) == 1:
                if messagebox.askyesno(
                    "确认", f'没有完全匹配的名称，是不是想选 "{candidates[0]}" ？'
                ):
                    name = candidates[0]
            elif len(candidates) > 1:
                messagebox.showwarning(
                    "提示",
                    "有多个相似的名称，请直接在上面的列表里勾选，而不要手动输入：\n"
                    + "\n".join(candidates),
                )
                return
            else:
                if not messagebox.askyesno(
                    "确认",
                    f'"{name}" 不在当前会话列表里，确定要按这个名字监听吗？'
                    "（必须和微信里显示的名称完全一致，否则后续会连接失败）",
                ):
                    return

        if name not in self.manual_names:
            self.manual_names.append(name)
            self.manual_label.config(text="已手动添加: " + ", ".join(self.manual_names))
        self.manual_entry.delete(0, tk.END)

    def _cancel(self):
        self.config_result = None
        self.destroy()

    def _next(self):
        chosen = [self.session_names[i] for i in self.listbox.curselection()]
        chosen += [n for n in self.manual_names if n not in chosen]
        if not chosen:
            messagebox.showwarning("提示", "至少选一个群/联系人")
            return

        chats = []
        for name in chosen:
            members = try_get_group_members(self.wx, name)
            picker = MemberPickerDialog(self, name, members)
            chats.append({"name": name, "senders": picker.result})

        self.config_result = {
            "chats": chats,
            "poll_interval_minutes": self.interval_var.get(),
            "save_media": {
                "savepic": self.media_var.get(),
                "savefile": self.media_var.get(),
                "savevoice": self.media_var.get(),
            },
        }
        self.destroy()


def main():
    root = tk.Tk()
    root.withdraw()
    lang_dialog = LanguageDialog(root)
    root.wait_window(lang_dialog)
    language = lang_dialog.result
    root.destroy()
    if not language:
        return

    wx = WeChat(language=language)
    sessions = wx.GetSessionList(reset=True)

    app = MainApp(wx, sessions)
    app.mainloop()

    if not app.config_result:
        print("已取消，未保存配置。")
        return

    # 摘要时间表、交易信号提醒等高级配置不在向导里问，重新配置时原样保留
    existing_config = {}
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            existing_config = json.load(f)

    config = {
        "language": language,
        **app.config_result,
        "summary_schedule": existing_config.get("summary_schedule")
        or DEFAULT_SUMMARY_SCHEDULE,
        "weekly_summary_schedule": existing_config.get("weekly_summary_schedule")
        or DEFAULT_WEEKLY_SCHEDULE,
    }
    for key in PRESERVED_ADVANCED_KEYS:
        if key in existing_config:
            config[key] = existing_config[key]
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)

    print(f"配置已保存到 {CONFIG_PATH}")
    print(f"监听对象: {[c['name'] for c in config['chats']]}")
    print(f"记录频率: 每 {config['poll_interval_minutes']} 分钟")
    print("摘要时间表: " + ", ".join(f"{e['label']}@{e['time']}" for e in config["summary_schedule"]))
    ws = config["weekly_summary_schedule"]
    print(f"周报时间表: 每周{ws['day']} {ws['time']}")
    print("\n接下来可以运行:")
    print("  python monitor.py --once     # 手动跑一次")
    print("  python monitor.py            # 常驻运行")


if __name__ == "__main__":
    main()
