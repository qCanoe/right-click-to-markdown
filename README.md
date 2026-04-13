# **Right Click to** MarkDown · Windows 右键一键转换为 Markdown文档

在资源管理器中右键文件，选择 **Convert to Markdown (MarkItDown)**，会在**同目录**生成与源文件同名的 `.md文件`。转换引擎来源于 [MarkItDown](https://github.com/microsoft/markitdown)；本仓库额外提供 `scripts/windows/` 中的右键菜单安装与无控制台调用封装。

---

## 一、环境要求

- **操作系统**：Windows 10/11（右键菜单仅适用于 Windows）
- **Python**：3.10 及以上（与 MarkItDown 官方要求一致）
- 建议用 **虚拟环境**，避免污染系统 Python（下面以 `venv` 为例，Anaconda 亦可）

---

## 二、安装 MarkItDown

**方式 A：从 PyPI 安装（常用）**

在 PowerShell 或「命令提示符」中执行：

```bash
python -m pip install --upgrade pip
python -m pip install "markitdown[all]"
```

若只需部分格式，可把 `[all]` 换成 `[pdf,docx,pptx]` 等组合。

**方式 B：从本仓库源码可编辑安装**

若你已克隆本仓库并希望用本地包开发/调试：

```bash
cd <本仓库根目录>
python -m pip install -e "packages/markitdown[all]"
```

**验证是否安装成功**（任选能成功的一条即可）：

```bash
markitdown --version
```

或：

```bash
python -m markitdown --version
```

若提示找不到命令，请确认安装时使用的 Python 对应目录下的 `Scripts` 已加入 **用户或系统 PATH**（资源管理器里右键转换时，PATH 可能比当前终端更「干净」，建议把该 `Scripts` 配进用户环境变量）。

---

## 三、安装本仓库的右键菜单脚本

1. **获取脚本**
  克隆本仓库，或至少保证本地有完整目录：  
   `scripts/windows/`（内含 `install-markitdown-context-menu.ps1`、`markitdown-convert.ps1`、`markitdown-convert-launcher.vbs`）。
2. **在仓库根目录打开 PowerShell**，执行：
  ```powershell
   Set-Location <本仓库根目录>
   .\scripts\windows\install-markitdown-context-menu.ps1
  ```
   若提示无法运行脚本（执行策略），可用：
3. **使菜单生效**
  若右键未立即出现新菜单项：在任务管理器中结束「Windows 资源管理器」再启动，或注销/重新登录。
4. **使用**
  在文件夹中右键目标文件 → **Convert to Markdown (MarkItDown)**。成功后同目录会出现 `原文件名.md`。

> **说明**：安装脚本会把 **绝对路径** 写入当前用户注册表。若你**移动了**本仓库所在文件夹，需在新位置**重新运行**一次安装命令。

---

## 四、卸载右键菜单

在仓库根目录执行：

```powershell
.\scripts\windows\install-markitdown-context-menu.ps1 -Uninstall
```

---

**仓库：** [github.com/qCanoe/right-click-to-markdown](https://github.com/qCanoe/right-click-to-markdown)

*转换核心与上游文档来自：[https://github.com/microsoft/markitdown](https://github.com/microsoft/markitdown)*