#!/bin/bash

# GAS 開發測試示範腳本
# 此腳本示範如何整合 clasp 與 Chrome DevTools MCP 進行開發

echo "╔════════════════════════════════════════════════╗"
echo "║  GAS 開發與測試工作流程示範                    ║"
echo "╚════════════════════════════════════════════════╝"
echo ""

# 顏色定義
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# 步驟 1: 檢查環境
echo -e "${BLUE}[步驟 1]${NC} 檢查開發環境..."
echo ""

if ! command -v clasp &> /dev/null; then
    echo -e "${YELLOW}⚠️  clasp 未安裝${NC}"
    echo "請執行: npm install -g @google/clasp"
    exit 1
fi

echo -e "${GREEN}✓${NC} clasp 已安裝: $(clasp --version 2>&1 | head -n 1)"

if [ ! -f ".clasp.json" ]; then
    echo -e "${YELLOW}⚠️  專案未連結${NC}"
    echo "請執行: clasp login && clasp clone <scriptId>"
    exit 1
fi

echo -e "${GREEN}✓${NC} 專案已連結"
echo ""

# 步驟 2: 推送程式碼
echo -e "${BLUE}[步驟 2]${NC} 推送程式碼至 GAS..."
echo ""

clasp push

if [ $? -eq 0 ]; then
    echo -e "${GREEN}✓${NC} 程式碼推送成功"
else
    echo -e "${YELLOW}⚠️  程式碼推送失敗，請檢查語法錯誤${NC}"
    exit 1
fi
echo ""

# 步驟 3: 開啟 Apps Script 編輯器
echo -e "${BLUE}[步驟 3]${NC} 開啟 Apps Script 編輯器..."
echo ""
echo "請在編輯器中執行以下任一函式進行測試："
echo ""
echo -e "${GREEN}快速測試：${NC}"
echo "  quickTest()          - 驗證基本功能"
echo ""
echo -e "${GREEN}完整測試：${NC}"
echo "  runAllTests()        - 執行所有測試套件"
echo "  testDomainModels()   - 只測試領域模型"
echo "  testExamService()    - 只測試 ExamService"
echo ""
echo -e "${GREEN}除錯工具：${NC}"
echo "  debugExamObject()    - 顯示 Exam 物件詳細資訊"
echo "  createDataSnapshot() - 建立資料快照"
echo ""
echo "按 Enter 開啟編輯器..."
read

clasp open

echo ""

# 步驟 4: 部署 Web App
echo -e "${BLUE}[步驟 4]${NC} 部署 Web App（選用）..."
echo ""
echo "如需視覺化測試介面，請在 Apps Script 編輯器中："
echo ""
echo "1. 點擊右上角「部署」→「新增部署作業」"
echo "2. 類型：選擇「網頁應用程式」"
echo "3. 執行身分：選擇「我」"
echo "4. 存取權：選擇「所有人」或「僅限自己」"
echo "5. 點擊「部署」"
echo "6. 複製 Web App URL"
echo ""
echo "Web App 功能："
echo "  - 📊 即時統計資訊"
echo "  - 🔄 重新載入資料"
echo "  - 💾 下載 JSON 快照"
echo "  - 📋 完整資料檢視"
echo ""
echo "按 Enter 繼續..."
read

# 步驟 5: Chrome DevTools 整合說明
echo ""
echo -e "${BLUE}[步驟 5]${NC} 使用 Chrome DevTools 進行除錯..."
echo ""
echo "當 Web App 開啟後，可以使用 Chrome DevTools："
echo ""
echo -e "${GREEN}基本除錯：${NC}"
echo "  1. 按 F12 開啟 DevTools"
echo "  2. 在 Console 中輸入："
echo "     currentData          - 查看當前資料"
echo "     loadData()           - 重新載入"
echo "     downloadJSON()       - 下載快照"
echo ""
echo -e "${GREEN}透過 MCP 自動化（VS Code）：${NC}"
echo "  - 在 VS Code 中可透過 MCP 控制瀏覽器"
echo "  - 擷取頁面快照"
echo "  - 執行 JavaScript 程式碼"
echo "  - 監控網路請求"
echo "  - 分析效能數據"
echo ""

# 步驟 6: 監視模式
echo -e "${BLUE}[步驟 6]${NC} 啟用監視模式（選用）..."
echo ""
echo "如需自動推送變更，可執行："
echo ""
echo "  npm run watch"
echo ""
echo "此模式會在檔案儲存時自動推送至 GAS"
echo "適合快速迭代開發"
echo ""
echo -e "${YELLOW}注意：${NC}確保程式碼無語法錯誤，否則推送會失敗"
echo ""
echo "是否啟動監視模式？(y/N)"
read -r response

if [[ "$response" =~ ^([yY][eE][sS]|[yY])$ ]]; then
    echo ""
    echo -e "${GREEN}啟動監視模式...${NC}"
    echo "按 Ctrl+C 停止"
    echo ""
    npm run watch
else
    echo ""
    echo -e "${BLUE}已略過監視模式${NC}"
fi

# 完成
echo ""
echo "╔════════════════════════════════════════════════╗"
echo "║  開發環境已就緒！                              ║"
echo "╚════════════════════════════════════════════════╝"
echo ""
echo "📚 延伸閱讀："
echo "  - QUICKSTART.md    - 快速開始指南"
echo "  - DEVELOPMENT.md   - 完整開發文件"
echo "  - AGENTS.md        - 專案結構規範"
echo ""
echo "🎯 下一步："
echo "  1. 在 GAS 編輯器執行 quickTest()"
echo "  2. 部署 Web App 進行視覺化測試"
echo "  3. 開始開發新功能！"
echo ""
echo "Happy Coding! 🚀"
echo ""
