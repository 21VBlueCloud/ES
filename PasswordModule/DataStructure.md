# 数据存储结构

文件使用通用的Json方式存储。

文件路径: "$home\MyPassword"
文件名："\<ComputerName>.json"

文件中包含多个加密对象。对象包含以下字段：

* DateTime：时间戳。用于标识该加密密码的创建或更新时间。
* Alias：凭据别名。唯一值。用于获取加密密码时方便使用。
* UserName：凭据用户名。
* EncryptedPwd：加密过的用户密码。
