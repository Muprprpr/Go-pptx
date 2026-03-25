# Development Guide / 开发指南

[English](#english) | [中文](#中文)

---

<a name="english"></a>

## English

### Local Development Setup

When developing locally, you need to use the local code instead of the remote repository. There are several ways to achieve this:

#### Method 1: Using `go.mod` replace directive (Recommended)

Add a `replace` directive to your `go.mod`:

```go
module github.com/Muprprpr/Go-pptx

go 1.24.5

// For local development, point to local directory
replace github.com/Muprprpr/Go-pptx => /path/to/your/local/Go-pptx
```

This tells Go to use the local directory instead of fetching from GitHub.

#### Method 2: Using `go.mod` with relative path

If your project using this library is in a sibling directory:

```go
replace github.com/Muprprpr/Go-pptx => ../Go-pptx
```

#### Method 3: Using `go.work` (Go 1.18+)

Create a `go.work` file in your workspace root:

```go
go 1.24.5

use (
    ./Go-pptx           // This library
    ./my-pptx-project   // Your project using this library
)
```

#### Method 4: Using GOPATH/src (Traditional)

```bash
# Clone to GOPATH
mkdir -p $GOPATH/src/github.com/Muprprpr
cd $GOPATH/src/github.com/Muprprpr
git clone https://github.com/Muprprpr/Go-pptx.git
```

### Quick Setup Script

Create a `dev_setup.sh` (or `dev_setup.bat` for Windows):

```bash
#!/bin/bash
# Run this in your project that uses go-pptx

# Add replace directive temporarily for development
go mod edit -replace github.com/Muprprpr/Go-pptx=/path/to/Go-pptx

# Verify
go list -m github.com/Muprprpr/Go-pptx
```

### Before Publishing

**Important:** Remove or comment out the `replace` directive before publishing your project:

```go
// replace github.com/Muprprpr/Go-pptx => ../Go-pptx  // Uncomment for local dev
```

Or use a separate `go.mod.dev` file and swap as needed.

---

<a name="中文"></a>

## 中文

### 本地开发设置

在本地开发时，需要使用本地代码而不是远程仓库。有几种方法可以实现：

#### 方法 1：使用 `go.mod` replace 指令（推荐）

在 `go.mod` 中添加 `replace` 指令：

```go
module github.com/Muprprpr/Go-pptx

go 1.24.5

// 本地开发时，指向本地目录
replace github.com/Muprprpr/Go-pptx => /path/to/your/local/Go-pptx
```

这告诉 Go 使用本地目录而不是从 GitHub 获取。

#### 方法 2：使用相对路径

如果你使用此库的项目在相邻目录：

```go
replace github.com/Muprprpr/Go-pptx => ../Go-pptx
```

#### 方法 3：使用 `go.work`（Go 1.18+）

在工作区根目录创建 `go.work` 文件：

```go
go 1.24.5

use (
    ./Go-pptx           // 本库
    ./my-pptx-project   // 使用本库的项目
)
```

#### 方法 4：使用 GOPATH/src（传统方式）

```bash
# 克隆到 GOPATH
mkdir -p $GOPATH/src/github.com/Muprprpr
cd $GOPATH/src/github.com/Muprprpr
git clone https://github.com/Muprprpr/Go-pptx.git
```

### Windows 路径示例

```go
// Windows 绝对路径
replace github.com/Muprprpr/Go-pptx => C:\Users\ASUS\mywork\Go-pptx

// Windows 相对路径
replace github.com/Muprprpr/Go-pptx => ..\Go-pptx
```

### 快速设置脚本

创建 `dev_setup.bat`（Windows）：

```batch
@echo off
REM 在使用 go-pptx 的项目中运行此脚本

REM 添加临时的 replace 指令用于开发
go mod edit -replace github.com/Muprprpr/Go-pptx=C:\Users\ASUS\mywork\Go-pptx

REM 验证
go list -m github.com/Muprprpr/Go-pptx
```

### 发布前

**重要：** 在发布项目之前，移除或注释掉 `replace` 指令：

```go
// replace github.com/Muprprpr/Go-pptx => ../Go-pptx  // 本地开发时取消注释
```

或使用单独的 `go.mod.dev` 文件，按需切换。

### 推荐的工作流程

```
项目结构示例：
/workspace
├── Go-pptx/                    # 本库（开发中）
│   ├── go.mod
│   ├── opc/
│   └── ...
├── my-pptx-app/                # 使用本库的应用
│   ├── go.mod                  # 包含 replace 指令
│   └── main.go
└── go.work                     # 可选：工作区文件
```

**go.work 示例：**
```go
go 1.24.5

use (
    ./Go-pptx
    ./my-pptx-app
)
```

使用 `go.work` 时，Go 会自动解析本地模块，无需 `replace` 指令。
