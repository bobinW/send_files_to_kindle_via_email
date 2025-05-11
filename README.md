# recruitment_system

# Kindle电子书发送工具

一个通过邮箱将EPUB格式电子书发送至Kindle设备的GUI应用程序，支持Gmail和QQ邮箱。

## 主要功能

- 📨 支持Gmail/QQ邮箱SMTP服务
- 📁 自动扫描指定目录的EPUB文件
- 🔒 安全保存邮箱配置（密码加密存储）
- 📊 带进度条的可视化发送过程
- ✅ 自动分类已发送/发送失败文件
- 📝 详细的日志记录（send_log.txt）

## 快速开始

1. **准备配置**
   - Gmail用户需启用[两步验证](https://myaccount.google.com/security)并创建[应用专用密码](https://myaccount.google.com/apppasswords)
   - QQ邮箱用户需获取[SMTP授权码](https://service.mail.qq.com/cgi-bin/help?subtype=1&&id=28&&no=1001256)

2. **目录结构**
   ```
   ├── Ebooks/              # 默认电子书目录
   │   ├── 已发送至Kindle/   # 成功发送的文件
   │   └── 发送失败/         # 发送失败的文件
   ├── send_log.txt         # 日志文件
   └── config.json          # 配置文件
   ```

3. **使用步骤**
   1. 选择电子书目录（默认为Ebooks）
   2. 选择邮箱服务商（Gmail/QQ）
   3. 输入邮箱账号和密码/授权码
   4. 填写Kindle接收邮箱（需在Amazon账户白名单中）
   5. 点击"发送到Kindle"

## 配置说明

### 邮箱设置
| 服务商 | SMTP服务器       | 端口 | 认证方式       |
|--------|------------------|------|----------------|
| Gmail  | smtp.gmail.com   | 587  | 应用专用密码    |
| QQ邮箱 | smtp.qq.com      | 465  | SMTP授权码      |

### 文件要求
- 仅支持.epub格式文件
- 单文件建议小于50MB
- 文件名避免包含特殊字符

## 注意事项

⚠️ **发送限制**
- Gmail每次最多发送300个文件
- QQ邮箱每次最多发送10个文件

🔧 **故障排查**
1. 发送失败时检查日志文件
2. 确认Kindle邮箱已添加到[Amazon认可发件人列表](https://www.amazon.cn/hz/mycd/myx#/home/settings/payment)
3. 网络连接异常时程序会自动重试3次
