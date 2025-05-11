import os
import smtplib
import logging
import time
import json
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.utils import formatdate, make_msgid
import hashlib
from email.header import Header
from email import charset

# 设置字符集，优先使用 quoted-printable 编码
charset.add_charset("utf-8", charset.QP, charset.QP, "utf-8")

# 修正获取 .exe 文件所在目录的逻辑
def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    else:
        return os.path.dirname(os.path.abspath(__file__))

# 获取 .exe 文件所在的目录
BASE_DIR = get_base_dir()

# 默认电子书目录（相对路径）
DEFAULT_EBOOKS_DIR = os.path.join(BASE_DIR, "Ebooks")

# 设置日志文件路径（相对路径）
LOG_FILE = os.path.join(BASE_DIR, "send_log.txt")

# 配置文件路径（相对路径）
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")

# 设置 SentToKindle 和 FailedToSend 子目录
def setup_directories(ebooks_dir):
    sent_dir = os.path.join(ebooks_dir, "已发送至Kindle")
    failed_dir = os.path.join(ebooks_dir, "发送失败")
    for dir_path in [ebooks_dir, sent_dir, failed_dir]:
        if not os.path.exists(dir_path):
            os.makedirs(dir_path)
    return sent_dir, failed_dir

# 日志配置
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

def load_config():
    """加载保存的配置，如果文件不存在或损坏，返回 None"""
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r") as f:
                return json.load(f)
    except Exception as e:
        logging.error(f"Failed to load config file: {e}")
    return None

def save_config(config):
    """保存配置到文件，并更新邮箱历史记录和密码"""
    try:
        existing_config = load_config() or {}
        email_history = existing_config.get("email_history", [])

        current_email = config.get("email_username", "")
        if current_email and current_email not in email_history:
            email_history.append(current_email)
        config["email_history"] = email_history

        email_provider = config.get("email_provider", "Gmail")
        if email_provider == "Gmail":
            config["gmail_username"] = config.get("email_username", "")
            config["gmail_password"] = config.get("email_password", "")
        else:
            config["qq_username"] = config.get("email_username", "")
            config["qq_password"] = config.get("email_password", "")

        if "gmail_username" not in config:
            config["gmail_username"] = existing_config.get("gmail_username", "")
        if "gmail_password" not in config:
            config["gmail_password"] = existing_config.get("gmail_password", "")
        if "qq_username" not in config:
            config["qq_username"] = existing_config.get("qq_username", "")
        if "qq_password" not in config:
            config["qq_password"] = existing_config.get("qq_password", "")

        with open(CONFIG_FILE, "w") as f:
            json.dump(config, f, indent=4)
        logging.info("Configuration saved successfully.")
    except PermissionError as e:
        error_msg = (
            f"Permission denied: Unable to write config file to {CONFIG_FILE}\n"
            f"Please move the .exe file to a directory where you have write permissions (e.g., your Desktop or Documents folder).\n"
            f"Error details: {e}"
        )
        logging.error(error_msg)
        messagebox.showerror("错误", error_msg)
        raise
    except Exception as e:
        logging.error(f"Failed to save config file: {e}")
        messagebox.showerror("错误", f"无法保存配置：{e}")
        raise

def clean_password(password):
    """清理密码，去除首尾空格和多余的内部空格"""
    if not password:
        return password
    # 去除首尾空格，并将多个连续空格替换为一个
    cleaned = " ".join(password.split())
    # 如果是 Gmail 应用专用密码，移除所有空格
    if len(cleaned.replace(" ", "")) == 16 and cleaned.count(" ") > 0:
        cleaned = cleaned.replace(" ", "")
        logging.info("Detected potential Gmail App Password with spaces, removed spaces.")
    return cleaned

def test_smtp_connection(email_provider, email_username, email_password, retries=3, delay=5):
    """测试 SMTP 连接，支持重试机制，并提供详细的错误信息"""
    # 清理密码
    email_password = clean_password(email_password)
    for attempt in range(retries):
        try:
            if email_provider == "Gmail":
                smtp_server = "smtp.gmail.com"
                smtp_port = 587
                server = smtplib.SMTP(smtp_server, smtp_port, timeout=60)
                server.ehlo()
                server.starttls()
                server.ehlo()
            else:
                smtp_server = "smtp.qq.com"
                smtp_port = 465
                server = smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=60)
                server.ehlo()
            
            server.login(email_username, email_password)
            server.quit()
            logging.info("SMTP connection test successful.")
            return True
        except smtplib.SMTPAuthenticationError as e:
            logging.error(f"SMTP authentication failed: {e}")
            if email_provider == "Gmail":
                messagebox.showerror(
                    "错误",
                    "Gmail SMTP 认证失败。\n"
                    "请确保：\n"
                    "1. 已启用两步验证（在 Google 账户设置中）。\n"
                    "2. 使用的是应用专用密码（App Password），而不是账户密码。\n"
                    "3. 复制应用专用密码时不要包含空格（正确格式为 16 位连续字符，例如 'xxxxxxxxxxxxxxxx'）。\n"
                    "如何生成应用专用密码：\n"
                    "- 登录 Google 账户 > 安全性 > 两步验证 > 应用专用密码 > 生成新密码。\n"
                    f"错误详情：{e}"
                )
            else:
                messagebox.showerror(
                    "错误",
                    "QQ 邮箱 SMTP 认证失败。\n"
                    "请确保：\n"
                    "1. 使用的是授权码，而不是账户密码。\n"
                    "2. 复制授权码时不要包含空格。\n"
                    "如何获取授权码：\n"
                    "- 登录 QQ 邮箱 > 设置 > 账户 > POP3/SMTP 服务 > 开启服务并生成授权码。\n"
                    f"错误详情：{e}"
                )
            return False
        except smtplib.SMTPConnectError as e:
            logging.error(f"SMTP connection error (attempt {attempt + 1}/{retries}): {e}")
            if attempt == retries - 1:
                messagebox.showerror(
                    "错误",
                    f"无法连接到 {email_provider} SMTP 服务器。\n"
                    f"服务器：{smtp_server}:{smtp_port}\n"
                    "请检查：\n"
                    "1. 网络连接是否正常（尝试运行 'telnet smtp.gmail.com 587'）。\n"
                    "2. 防火墙或 VPN 是否阻止了连接。\n"
                    "3. DNS 是否能正确解析（尝试运行 'ping smtp.gmail.com'）。\n"
                    f"错误详情：{e}"
                )
                return False
            time.sleep(delay)
        except smtplib.SMTPServerDisconnected as e:
            logging.error(f"SMTP server disconnected (attempt {attempt + 1}/{retries}): {e}")
            if attempt == retries - 1:
                messagebox.showerror(
                    "错误",
                    f"{email_provider} SMTP 服务器断开连接。\n"
                    "请检查：\n"
                    "1. 网络是否稳定。\n"
                    "2. SMTP 配置是否正确。\n"
                    f"错误详情：{e}"
                )
                return False
            time.sleep(delay)
        except Exception as e:
            logging.error(f"SMTP connection failed (attempt {attempt + 1}/{retries}): {e}")
            if attempt == retries - 1:
                messagebox.showerror(
                    "错误",
                    f"SMTP 连接失败。\n"
                    f"请检查网络、用户名和密码/授权码是否正确。\n"
                    f"错误详情：{e}"
                )
                return False
            time.sleep(delay)
    return False

def is_valid_epub(epub_path):
    """验证文件是否为有效的 EPUB 文件（检查文件是否以 ZIP 格式开头并验证文件完整性）"""
    try:
        with open(epub_path, "rb") as f:
            magic = f.read(2)
            if magic != b"PK":
                return False
            f.seek(0)
            hasher = hashlib.sha256()
            while chunk := f.read(8192):
                hasher.update(chunk)
            file_hash = hasher.hexdigest()
            logging.info(f"File {epub_path} hash: {file_hash}")
        return True
    except Exception as e:
        logging.error(f"Failed to validate EPUB file {epub_path}: {e}")
        return False

def send_to_kindle(epub_path, email_provider, email_username, email_password, kindle_email, retries=3, delay=5):
    """发送电子书到 Kindle，支持重试机制"""
    # 清理密码
    email_password = clean_password(email_password)
    for attempt in range(retries):
        try:
            if not is_valid_epub(epub_path):
                logging.error(f"Invalid EPUB file: {epub_path}")
                return False

            filename = os.path.basename(epub_path)
            book_name = os.path.splitext(filename)[0]
            logging.info(f"Extracted book name: {book_name} from file: {filename}")

            msg = MIMEMultipart()
            msg["From"] = email_username
            msg["To"] = kindle_email
            # 设置邮件主题为书名（去掉扩展名后的文件名）
            subject_header = Header(book_name, "utf-8")
            msg["Subject"] = subject_header.encode(maxlinelen=76, linesep="\r\n")
            msg["Date"] = formatdate(localtime=True)
            msg["Message-ID"] = make_msgid()

            with open(epub_path, "rb") as attachment:
                part = MIMEApplication(
                    attachment.read(),
                    _subtype="epub+zip"
                )

            # 使用 Header 明确指定 quoted-printable 编码
            filename_header = Header(filename, "utf-8", header_name="Content-Disposition")
            filename_encoded = filename_header.encode(maxlinelen=76, linesep="\r\n")

            # 构造 Content-Disposition 头
            content_disposition = f"attachment; filename=\"{filename_encoded}\""
            part.add_header(
                "Content-Disposition",
                content_disposition
            )
            part.add_header(
                "Content-Type",
                "application/epub+zip"
            )
            logging.debug(f"Content-Disposition header: {content_disposition}")

            msg.attach(part)

            logging.debug(f"Email content:\n{msg.as_string()}")

            if email_provider == "Gmail":
                smtp_server = "smtp.gmail.com"
                smtp_port = 587
                server = smtplib.SMTP(smtp_server, smtp_port, timeout=60)
                server.starttls()
            else:
                smtp_server = "smtp.qq.com"
                smtp_port = 465
                server = smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=60)

            server.login(email_username, email_password)
            server.sendmail(email_username, kindle_email, msg.as_string())
            server.quit()
            logging.info(f"Sent: {epub_path} -> {kindle_email}")
            return True
        except Exception as e:
            logging.error(f"Failed to send {epub_path} (attempt {attempt + 1}/{retries}): {e}")
            if attempt == retries - 1:
                return False
            time.sleep(delay)
    return False

def move_file(epub_path, success, sent_dir, failed_dir):
    """移动文件到成功或失败目录"""
    dest_dir = sent_dir if success else failed_dir
    base_name = os.path.basename(epub_path)
    dest_path = os.path.join(dest_dir, base_name)
    counter = 1
    while os.path.exists(dest_path):
        name, ext = os.path.splitext(base_name)
        dest_path = os.path.join(dest_dir, f"{name}_{counter}{ext}")
        counter += 1
    os.rename(epub_path, dest_path)
    logging.info(f"Moved: {epub_path} -> {dest_path}")

class KindleSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("发送文件到Kindle")
        # 调整界面尺寸为 1024x640
        self.root.geometry("1024x640")
        self.root.configure(bg="#B0E0E6")

        self.config = load_config()
        self.ebooks_dir = tk.StringVar()
        self.email_username = tk.StringVar()
        self.email_password = tk.StringVar()
        self.kindle_email = tk.StringVar()
        self.progress = tk.DoubleVar()
        self.connection_status = tk.StringVar(value="未连接")
        self.email_provider = tk.StringVar(value="Gmail")
        self.email_history = []
        self.gmail_password = ""
        self.qq_password = ""

        if self.config:
            self.ebooks_dir.set(self.config.get("ebooks_dir", DEFAULT_EBOOKS_DIR))
            self.email_username.set(self.config.get("email_username", ""))
            self.email_password.set(self.config.get("email_password", ""))
            self.kindle_email.set(self.config.get("kindle_email", ""))
            self.email_provider.set(self.config.get("email_provider", "Gmail"))
            self.email_history = self.config.get("email_history", [])
            self.gmail_password = self.config.get("gmail_password", "")
            self.qq_password = self.config.get("qq_password", "")
        else:
            self.ebooks_dir.set(DEFAULT_EBOOKS_DIR)

        self.style = ttk.Style()
        self.style.configure("TLabel", background="#B0E0E6", foreground="#191970", font=("微软雅黑", 14))
        self.style.configure("TEntry", fieldbackground="#FFFFFF", foreground="#191970", font=("微软雅黑", 14))
        self.style.configure("TButton", background="#87CEFA", foreground="#191970", font=("微软雅黑", 14, "bold"))
        self.style.configure("TProgressbar", troughcolor="#E0FFFF", background="#4682B4")
        self.style.configure("TCombobox", fieldbackground="#FFFFFF", foreground="#191970", font=("微软雅黑", 14))

        self.create_gui()

    def create_gui(self):
        tk.Label(self.root, text="发送电子书到Kindle", font=("微软雅黑", 22, "bold"), bg="#B0E0E6", fg="#191970").pack(pady=5)

        tk.Label(self.root, text="电子书目录：", bg="#B0E0E6", fg="#191970", font=("微软雅黑", 14)).pack(pady=2)
        ttk.Entry(self.root, textvariable=self.ebooks_dir, width=70).pack()
        ttk.Button(self.root, text="浏览", command=self.browse_directory).pack(pady=2)

        tk.Label(self.root, text="邮箱提供商：", bg="#B0E0E6", fg="#191970", font=("微软雅黑", 14)).pack(pady=2)
        provider_frame = tk.Frame(self.root, bg="#B0E0E6")
        provider_frame.pack()
        ttk.Radiobutton(provider_frame, text="Gmail", value="Gmail", variable=self.email_provider, style="TLabel").pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(provider_frame, text="QQ邮箱", value="QQ", variable=self.email_provider, style="TLabel").pack(side=tk.LEFT, padx=10)

        tk.Label(self.root, text="邮箱用户名：", bg="#B0E0E6", fg="#191970", font=("微软雅黑", 14)).pack(pady=2)
        self.email_combobox = ttk.Combobox(self.root, textvariable=self.email_username, width=67, values=self.email_history, style="TCombobox")
        self.email_combobox.pack()
        self.email_combobox.bind("<<ComboboxSelected>>", self.update_email_provider_and_password)
        self.email_combobox.bind("<KeyRelease>", self.update_email_provider_and_password)

        tk.Label(self.root, text="邮箱密码：", bg="#B0E0E6", fg="#191970", font=("微软雅黑", 14)).pack(pady=2)
        ttk.Entry(self.root, textvariable=self.email_password, show="*", width=70).pack()

        tk.Label(self.root, text="提示：Gmail需使用应用专用密码（需启用两步验证）。QQ邮箱需使用授权码。", 
                 bg="#B0E0E6", fg="#191970", font=("微软雅黑", 14)).pack(pady=2)

        tk.Label(self.root, text="提示：Gmail一次最多发送300个文件，QQ邮箱一次最多发送10个文件。", 
                 bg="#B0E0E6", fg="#191970", font=("微软雅黑", 14)).pack(pady=2)

        tk.Label(self.root, text="Kindle邮箱：", bg="#B0E0E6", fg="#191970", font=("微软雅黑", 14)).pack(pady=2)
        ttk.Entry(self.root, textvariable=self.kindle_email, width=70).pack()

        tk.Label(self.root, text="连接状态：", bg="#B0E0E6", fg="#191970", font=("微软雅黑", 14)).pack(pady=2)
        tk.Label(self.root, textvariable=self.connection_status, bg="#B0E0E6", fg="#191970", font=("微软雅黑", 14, "italic")).pack()

        tk.Label(self.root, text="发送进度：", bg="#B0E0E6", fg="#191970", font=("微软雅黑", 14)).pack(pady=2)
        self.progress_bar = ttk.Progressbar(self.root, variable=self.progress, maximum=100, length=800, style="TProgressbar")
        self.progress_bar.pack()

        button_frame = tk.Frame(self.root, bg="#B0E0E6")
        button_frame.pack(pady=5)

        self.send_button = ttk.Button(button_frame, text="发送到Kindle", command=self.start_sending)
        self.send_button.pack(side=tk.LEFT, padx=15)

        self.finish_button = ttk.Button(button_frame, text="完成", command=self.root.quit, state=tk.DISABLED)
        self.finish_button.pack(side=tk.LEFT, padx=15)

    def update_email_provider_and_password(self, event=None):
        email = self.email_username.get().strip().lower()
        if email.endswith("@gmail.com"):
            self.email_provider.set("Gmail")
            self.email_password.set(self.gmail_password)
        elif email.endswith("@qq.com"):
            self.email_provider.set("QQ")
            self.email_password.set(self.qq_password)
        else:
            if self.email_provider.get() == "Gmail":
                self.email_password.set(self.gmail_password)
            else:
                self.email_password.set(self.qq_password)

    def browse_directory(self):
        directory = filedialog.askdirectory(initialdir=self.ebooks_dir.get())
        if directory:
            self.ebooks_dir.set(directory)

    def start_sending(self):
        self.send_button.config(state=tk.DISABLED)
        self.finish_button.config(state=tk.DISABLED)
        self.progress.set(0)
        self.connection_status.set("连接中...")
        self.root.update()

        ebooks_dir = self.ebooks_dir.get()
        email_provider = self.email_provider.get()
        email_username = self.email_username.get()
        email_password = self.email_password.get()
        kindle_email = self.kindle_email.get()

        if not ebooks_dir or not email_username or not email_password or not kindle_email:
            messagebox.showerror("错误", "请填写所有字段！")
            self.connection_status.set("未连接")
            self.send_button.config(state=tk.NORMAL)
            return

        files = [f for f in os.listdir(ebooks_dir) if f.lower().endswith(".epub")]
        if not files:
            messagebox.showinfo("提示", "目录中没有找到EPUB文件（仅支持 .epub 扩展名）。")
            logging.info("No EPUB files found in the directory.")
            self.send_button.config(state=tk.NORMAL)
            self.finish_button.config(state=tk.NORMAL)
            return

        file_count = len(files)
        max_files = 300 if email_provider == "Gmail" else 10

        if email_provider == "Gmail" and file_count > max_files:
            messagebox.showerror("错误", f"文件数量过多！Gmail一次最多发送{max_files}个文件，您有{file_count}个文件。")
            self.connection_status.set("未连接")
            self.send_button.config(state=tk.NORMAL)
            return

        if email_provider == "QQ" and file_count > max_files:
            messagebox.showinfo("提示", "QQ邮箱一次最多发送10个文件，将发送前10个文件。")
            files = files[:max_files]
            file_count = len(files)

        config = {
            "ebooks_dir": ebooks_dir,
            "email_provider": email_provider,
            "email_username": email_username,
            "email_password": email_password,
            "kindle_email": kindle_email,
        }
        try:
            save_config(config)
        except Exception:
            self.connection_status.set("未连接")
            self.send_button.config(state=tk.NORMAL)
            return

        if not test_smtp_connection(email_provider, email_username, email_password):
            self.connection_status.set("连接失败")
            self.send_button.config(state=tk.NORMAL)
            return
        self.connection_status.set("连接成功")
        self.root.update()

        sent_dir, failed_dir = setup_directories(ebooks_dir)

        total_files = len(files)
        for i, file in enumerate(files):
            if file == os.path.basename(LOG_FILE):
                continue
            file_path = os.path.join(ebooks_dir, file)
            file_size = os.path.getsize(file_path) / (1024 * 1024)
            logging.info(f"Sending: {file}, Size: {file_size:.2f} MB")
            success = False
            for attempt in range(3):
                if send_to_kindle(file_path, email_provider, email_username, email_password, kindle_email):
                    success = True
                    break
                else:
                    logging.warning(f"Failed to send {file} (attempt {attempt + 1}/3). Retrying...")
                    time.sleep(30)
            move_file(file_path, success, sent_dir, failed_dir)
            self.progress.set((i + 1) / total_files * 100)
            self.root.update()

        messagebox.showinfo("成功", "处理完成！请检查“已发送至Kindle”和“发送失败”文件夹。")
        self.send_button.config(state=tk.NORMAL)
        self.finish_button.config(state=tk.NORMAL)

def main():
    root = tk.Tk()
    app = KindleSenderApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()