# koishi-plugin-lostark-schedule

[![npm](https://img.shields.io/npm/v/koishi-plugin-lostark-schedule?style=flat-square)](https://www.npmjs.com/package/koishi-plugin-lostark-schedule)

## 命运方舟周本排期插件

### 安装

通过 `Koishi 插件市场` 安装

### 使用

- 配置地图信息 **必填**
  - 在插件配置页面中, 进行对应配置操作.

- 配置 `Smtp` 邮箱信息 **必填**
  - 「account」对应你的邮箱账号，「password」对应你的授权码
  - 「smtp」为发送服务器, 需要填写对应的「host」和「port」
  - 不同邮箱服务获取授权码的方式也有所不同，可以参考下面的主流邮件服务进行配置
  - **QQ 邮箱**
    1. 发送服务器：`smtp.qq.com`, 端口号 `465` 或 `587`
    2. 参考：[什么是授权码，它又是如何设置？](https://service.mail.qq.com/detail/0/75)
  - **网易 163 邮箱**
    1. 发送服务器：`smtp.163.com`, 端口号 `465` 或 `994`
    2. 参考：[网易邮箱 IMAP 服务](https://mail.163.com/html/110127_imap/index.htm)
  - **Outlook**
    1. 发送服务器：`smtp-mail.outlook.com`, 端口号 `587`
    2. 参考：[Outlook.com 的 POP、IMAP 和 SMTP 设置](https://support.microsoft.com/zh-cn/office/outlook-com-%E7%9A%84-pop-imap-%E5%92%8C-smtp-%E8%AE%BE%E7%BD%AE-d088b986-291d-42b8-9564-9c414e2aa040)
  - **Gmail**
    1. 发送服务器：`smtp.gmail.com`, 端口号 `465`
    2. 参考：[通过其他电子邮件平台查看 Gmail](https://support.google.com/mail/answer/7126229?hl=zh-Hans#zippy=%2C%E7%AC%AC-%E6%AD%A5%E6%A3%80%E6%9F%A5-imap-%E6%98%AF%E5%90%A6%E5%B7%B2%E5%90%AF%E7%94%A8%2C%E7%AC%AC-%E6%AD%A5%E5%9C%A8%E7%94%B5%E5%AD%90%E9%82%AE%E4%BB%B6%E5%AE%A2%E6%88%B7%E7%AB%AF%E4%B8%AD%E6%9B%B4%E6%94%B9-smtp-%E5%92%8C%E5%85%B6%E4%BB%96%E8%AE%BE%E7%BD%AE)

### 指令

- `副本名称 1大2小3奶 备注信息`
  - 其中副本名称为你在配置页面配置的副本名称: 如 `日月鹿 8` 就表示 `日月鹿`
  - 例如：`日月鹿 1大2小3奶`
  - 例如：`日月鹿 1大2小3奶 周三不在`
  - 例如：`日月鹿 @某人 1大2小3奶 周三不在` 这里可以代替某人代报名参加

- `副本名称列表`
  - 列出此副本下一期参加的人员所有信息
  - 例如：`日月鹿列表` 将返回所有下一期参加的人员信息以及合计人数信息

- `副本名称查询 @某人`
  - 查询某人在此副本下一期的参加信息
  - 例如：`日月鹿查询 @某人` 将返回某人在此副本下一期的参加信息

- `导出 副本名称`
  - 选项 `-l` 导出此副本上一期排班 Excel 表格
  - 导出此副本下一期排班 Excel 表格
  - 例如：`导出 日月鹿` 当前将通过邮件发送给触发指令的人员邮箱
