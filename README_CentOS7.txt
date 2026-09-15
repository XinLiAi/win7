============================================================
 设备密码管理系统 v3.0 多用户权限版
 CentOS 7 安装与运行说明
============================================================

一、角色与权限
------------------------------------------------------------
  管理员(admin)   : 管理用户（创建/编辑/删除/重置密码，三种角色
                    账号只能由管理员创建） + 查看【用户登录日志】
                    不参与设备增删改查
  操作员(operator): 设备增删改查 + Excel 导入导出
                    不能创建账号、不能看日志
  审计员(auditor) : 只读查看【操作日志】（设备增删改查等）
                    + 导出日志（CSV / Excel）
                    无任何写操作

  两类日志严格分离：
    登录日志  -> 谁在何时登录/退出/失败 -> 管理员可见
    操作日志  -> 设备增删改查、用户管理等 -> 审计员可见

二、环境要求
------------------------------------------------------------
  * CentOS 7 / RHEL 7（或其他 Linux 发行版）
  * Python 3.6+，tkinter，sqlite3
  * openpyxl（可选，仅 Excel 导入导出 / 导出Excel日志需要）

三、CentOS 7 一键安装
------------------------------------------------------------
  sudo yum install -y python3 python3-tkinter python3-pip
  sudo yum install -y wqy-microhei-fonts wqy-zenhei-fonts   # 中文字体
  sudo pip3 install openpyxl                                  # 可选

四、运行
------------------------------------------------------------
  chmod +x run.sh
  ./run.sh
  # 或直接: python3 device_manager.py

五、使用流程
------------------------------------------------------------
  1) 首次运行【强制】创建管理员账号（用户名+密码，密码≥8位含字母数字）
  2) 之后每次启动，用 用户名+密码 登录（失败3次自动退出）
  3) 登录后按角色显示对应功能：
     - 管理员：用户管理（添加/编辑/删除/重置密码，可设三种角色）
              + 登录日志（刷新/清空）
     - 操作员：设备管理（增删改查/导入导出/搜索/明文切换）
     - 审计员：操作日志（刷新/导出CSV/导出Excel）

  保护规则：
    - 不能删除/禁用当前登录的账号
    - 系统必须保留至少一个启用的管理员（最后一个管理员不可删/禁用/改角色）
    - 操作员账号、审计员账号只能由管理员创建

六、数据与安全说明
------------------------------------------------------------
  数据目录: ~/.device_manager/
    devices.db      设备数据 + 用户表 + 两类日志（SQLite）
    secret.key      设备密码加密密钥（文件权限 600）

  安全特性：
    * 登录密码 PBKDF2-HMAC-SHA256 加盐哈希（12万次迭代）
    * 设备密码 CTR 流加密存储，列表默认脱敏（******）
    * 全程审计：登录成功/失败/退出、设备增删改查、导入导出、用户管理
    * 操作员设备操作自动写入操作日志供审计员追溯

  重要提醒：
    * 请备份 ~/.device_manager/secret.key，丢失后已存设备密码无法解密
    * 忘记登录密码：删除 ~/.device_manager/password.hash 无效（v3 已改为用户表）。
      重置办法：数据库 devices.db 中 users 表，或重新初始化（会丢数据）。
      更稳妥：联系管理员用"重置密码"功能直接重设。

七、故障排查
------------------------------------------------------------
  1) 中文显示方块   -> 安装 wqy-microhei-fonts / wqy-zenhei-fonts
  2) 提示缺 tkinter -> sudo yum install -y python3-tkinter
  3) Excel 不可用   -> sudo pip3 install openpyxl
  4) 远程无界面     -> VNC，或 ssh -X，或 export DISPLAY=:0

八、打包为 Windows EXE + 一键安装包（可选）
------------------------------------------------------------
  用 PyInstaller 打成单个 .exe（已内嵌 Python 运行时和所有依赖，
  目标电脑无需安装 Python），再用 Inno Setup 包成一键安装包
  （带安装向导、桌面快捷方式、开始菜单、卸载入口）。

  注意：必须在 Windows 环境上打包；Linux 只能打出 Linux 可执行文件。

  方式一（推荐·无需本机 Windows）: GitHub Actions 云构建
    用 GitHub 免费提供的 Windows 构建机自动打包 EXE + 安装包：
    1. 把项目文件上传到一个 GitHub 仓库
       （device_manager.py / device_manager.spec / device_manager.ico /
        device_manager_installer.iss / build_exe.bat / README_CentOS7.txt）
    2. 把 .github/workflows/build_exe.yml 也放进去
    3. 仓库页 -> Actions -> "Build Windows EXE + Installer" -> Run workflow
    4. 等 3~5 分钟 -> 打开本次运行 -> Artifacts -> 下载
       DeviceManager-package.zip（内含 DeviceManager.exe 和安装包）

  方式二（本机有 Windows）: 双击 build_exe.bat 一键打包
    前提：安装 Python 3.8+（勾选 Add to PATH）
    可选：安装 Inno Setup 6（https://jrsoftware.org/isdl.php），
          装了就会同时生成一键安装包，没装只生成绿色版 EXE
    产物:
      dist\DeviceManager.exe                     绿色版（拷走就能用）
      Output\设备密码管理系统_Setup_v3.1.exe      一键安装包

  方式三（手动命令）:
    pip install pyinstaller openpyxl
    python -m PyInstaller --noconfirm --clean device_manager.spec
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" device_manager_installer.iss

  说明：
    * 单文件 EXE 已内嵌 Python + openpyxl + tkinter，目标电脑零依赖
    * 安装包会把程序装到 C:\Program Files\DeviceManager\，创建快捷方式和卸载入口
    * 数据保存在 用户目录\.device_manager\ 下（卸载时默认保留，不丢数据）
    * 杀软可能误报，可加白名单；首次运行会较慢（解压到临时目录）
