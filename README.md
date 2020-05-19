这是一个可以将 ACCESS 数据库文件(.mdb)导入 MySQL 数据库中的简单脚本。使用了 [pyodbc](https://github.com/mkleehammer/pyodbc) 来实现对 .mdb 文件的读取。

**使用方法如下：**

* 下载 mdb2sql.py 文件；
* 在你的项目中加入`from mdb2sql import mdb2sql`；
* 根据提示输入`mdb2sql()`函数所需的参数，运行完成即可。

> 注意：
> * 确保你的电脑中有 ACCESS 的驱动程序，具体细节可以参考[这里](https://github.com/mkleehammer/pyodbc/wiki/Connecting-to-Microsoft-Access)；
> * filepath 参数必须输入**绝对路径**！
