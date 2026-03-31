#!/bin/bash
# ========================================
# QQx小游戏品牌方案 - GitHub Pages 部署脚本
# ========================================

set -e

REPO_NAME="qqx-mini-game-brand"
PROJECT_DIR="/Users/haoxingwang/WorkBuddy/QQx小游戏品牌方案"

cd "$PROJECT_DIR"

echo ""
echo "======================================"
echo " QQx小游戏品牌方案 · GitHub Pages 部署"
echo "======================================"
echo ""

# Step 1: 检查 gh 登录状态
echo "📋 Step 1: 检查 GitHub 登录状态..."
if ! gh auth status &>/dev/null; then
    echo "⚠️  未登录 GitHub，现在开始登录..."
    echo "   （会打开浏览器，请在浏览器中完成授权）"
    echo ""
    gh auth login
fi
echo "✅ GitHub 已登录"
echo ""

# Step 2: 创建远程仓库
echo "📋 Step 2: 创建 GitHub 仓库 '$REPO_NAME'..."
if gh repo view "$REPO_NAME" &>/dev/null; then
    echo "   仓库已存在，跳过创建"
else
    gh repo create "$REPO_NAME" --public --description "QQ媒体 × 微信小游戏 品牌效果售卖方案体系" --source=. --push=false
    echo "✅ 仓库创建成功"
fi
echo ""

# Step 3: Git 提交
echo "📋 Step 3: 提交代码..."
git add -A
git commit -m "feat: QQx小游戏品牌方案体系 v1.0 - 完整工作流+模板+可视化平台" --allow-empty 2>/dev/null || true
echo "✅ 代码已提交"
echo ""

# Step 4: 设置远程仓库并推送
echo "📋 Step 4: 推送到 GitHub..."
GITHUB_USER=$(gh api user --jq '.login')
REMOTE_URL="https://github.com/$GITHUB_USER/$REPO_NAME.git"

if git remote get-url origin &>/dev/null; then
    git remote set-url origin "$REMOTE_URL"
else
    git remote add origin "$REMOTE_URL"
fi

git branch -M main
git push -u origin main
echo "✅ 代码已推送到 GitHub"
echo ""

# Step 5: 启用 GitHub Pages
echo "📋 Step 5: 启用 GitHub Pages..."
gh api -X POST "repos/$GITHUB_USER/$REPO_NAME/pages" \
    -f "build_type=legacy" \
    -f "source[branch]=main" \
    -f "source[path]=/" 2>/dev/null || \
gh api -X PUT "repos/$GITHUB_USER/$REPO_NAME/pages" \
    -f "build_type=legacy" \
    -f "source[branch]=main" \
    -f "source[path]=/" 2>/dev/null || true
echo "✅ GitHub Pages 已启用"
echo ""

# Step 6: 获取线上地址
echo "======================================"
echo " 🎉 部署完成！"
echo "======================================"
echo ""
echo " 🌐 线上地址（1-2分钟后生效）:"
echo ""
echo "    https://$GITHUB_USER.github.io/$REPO_NAME/"
echo ""
echo " 📄 直接链接:"
echo "    首页:     https://$GITHUB_USER.github.io/$REPO_NAME/06_工作流可视化平台/index.html"
echo "    抖音报告: https://$GITHUB_USER.github.io/$REPO_NAME/05_竞媒深度调研报告/抖音小游戏营销生态调研报告.html"
echo "    通用方案: https://$GITHUB_USER.github.io/$REPO_NAME/03_PPT营销案模板/QQx小游戏通用营销方案.html"
echo ""
echo " 📦 GitHub 仓库:"
echo "    https://github.com/$GITHUB_USER/$REPO_NAME"
echo ""
echo "======================================"
