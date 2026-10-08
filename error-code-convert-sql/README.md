# error-code-convert-sql

把 Excel 错误码表转换为 `internationalization` 表的 SQL 文件。

## Excel 格式

工具会自动寻找包含以下表头的行，数据从表头下一行开始读取：

| 传参值 | key | 中文 | 英文 |
| --- | --- | --- | --- |
| powerStartupFailure | 1 | 电源启动失败 | Power startup failure |

生成规则：

- `i18n_key` = `RWI.` + `传参值`
- 中文 `i18n_value` = `key` + 空格 + `中文`
- 英文 `i18n_value` = `key` + 空格 + `英文`
- 输出包含一段 `DELETE`、一段中文 `INSERT`、一段英文 `INSERT`

## 构建（推荐，无需下载依赖）

```powershell
.\build.ps1
```

这个版本使用纯 Java（JDK 25 编译）解析 `.xlsx`，不需要访问 Maven 仓库。

## Maven 构建

```powershell
mvn clean package
```

如果本机 Maven 仓库或公司内网仓库不可用，请使用上面的 `build.ps1`。

## 使用

### 双击启动（图形界面）

双击 `启动工具.bat`，弹出窗口后用「浏览...」选择 Excel 文件和输出 SQL 路径，点「生成 SQL」即可。

（首次运行若无 JAR 会自动调用 `build.ps1` 编译；需本机已安装 Java 并配置 PATH。）

### 命令行

```powershell
java -jar target\error-code-convert-sql-1.0.0.jar C:\Users\Reid.Liu\Downloads\ErrorCode-临工.xlsx
```

默认会在当前目录生成 `errorcode.sql`。

也可以指定输出文件：

```powershell
java -jar target\error-code-convert-sql-1.0.0.jar C:\Users\Reid.Liu\Downloads\ErrorCode-临工.xlsx D:\workspace\myself\error-code-convert-sql\errorcode.sql
```

常用可选参数：

```powershell
java -jar target\error-code-convert-sql-1.0.0.jar <excel文件> `
  --output errorcode.sql `
  --sheet Sheet1 `
  --prefix RWI. `
  --timestamp "2021-11-26 15:30:09.301" `
  --create-by kuka `
  --create-app OptionalCollection:312 `
  --group robot
```
