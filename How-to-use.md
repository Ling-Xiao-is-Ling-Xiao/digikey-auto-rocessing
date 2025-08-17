# 怎么使用这个项目？

## 一. 搭建运行环境
一. 下载python编译器

1. 你需要一个python编译器，可以从[python官网](https://www.python.org/)官网上下载。

2. 下载完成之后，按照提示完成安装，_**安装路径不能包含非英文字符和空格**_。

3. 安装完成后，需要完成配置，win10和win11的操作略有差异。

* win10: 右键“我的电脑”，选择“属性”，“高级系统设置”，选择“环境变量”，把第一步安装的python编辑器安装路径添加到“PATH”变量里。
* win11: 右键“我的电脑”，选择“属性”，在弹出来的设置里向下滑在“设备规格”的“相关链接”里找到“高级系统设置”，然后选择"环境变量"，把第一步安装的python编辑器安装路径添加到“PATH”变量里。

二. 下载vscode

> 当然，直接用pythonIDE运行这个项目也行，这是非必要步骤。

你可以从[vscode官网](https://code.visualstudio.com/Download)，然后根据提示完成安装。

三. 下载项目文件并且配置环境

1. 在项目的"code"页面，点击绿色的"code"按钮，选择"Loacl"，再点击"Download Zip"按钮，下载项目zip文件。

2. 把项目zip文件解压到一个安全（下载目录，C盘，桌面，回收站之外的不含有非英文字符的）路径，然后删除除了.py和.json之外的无用文件。

3. 按住"win+R"打开运行，输入"cmd"，输入"cd + 你解压的地址"，然后输入以下内容并且回车。
* pip install requests
* pip install openpyxl

4. 使用pythonIDE或者vscode打开excel_to_digikey.py。

## 二.申请digikeyAPI

1. 打开[digikeyAPI官网](https://www.digikey.com/api)。

2. 点击页面中下方的"启动开发者门户"。

3. 点击右上角"注册或登录"，如果有digikry账号就登录，否则就按照提示注册。

4. 点击右上角的“组织”，然后点击“+creatr organization”，随便输入一个名字,点击绿色按钮。

5. 点击“Production Apps”选项卡，然后点击“+creatr production apps”。

6. "Production App name","OAuth Callback","Description"按照实际情况填写，"Select one or more Production products:"处需要勾选“Product Information V4“，然后点击最下面的蓝色按钮。

7. 然后点击自己的项目名，进入详情，在”Credentials“处的”Client ID“，”Client Secret“点击”Shou Key“，然后复制。

8. 使用pythonIDE或者vscode打开digikey.py，在"DIGIKEY_CLIENT_ID"处输入刚刚的”Client ID“，在"DIGIKEY_CLIENT_SECRET"出输入”Client Secret“。
> 注意，不能换行。

9. 按下"ctrl+s"保存，然后运行。
