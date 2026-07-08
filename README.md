# 秀朗國小特色與學習共同體課例探究專案

本專案旨在展示 **新北市永和區秀朗國民小學 (秀朗國小)** 的校史與現代特色，並包含針對國小一年級「時鐘素養題」公開課影片的學習共同體 (SLC) 課例探究分析成果。

## 專案內容與產出

### 1. 秀朗國小學校特色
- 包含世界紀錄歷史榮耀、多元藝文活動（管弦樂團與歌仔戲團）、體育基石（游泳隊）及現代化精緻與科技教育轉型。
- 生成的 Word 簡報文件：`秀朗國小特色簡報內容.docx`
- 生成的 Python 建立腳本：`create_presentation_word.py`

### 2. 學習共同體 (SLC) 課例探究
- 針對公開課影片（1150102一年級時鐘素養題_挑戰題）進行語音轉譯、逐字稿分析，並透過「描述-詮釋-反思」三維度進行分析。
- 轉譯逐字稿與字幕：`output/transcript.txt`、`output/subtitles.srt`
- 學習共同體課例分析報告：`output/analysis.txt`
- 16:9 簡報投影片 (PowerPoint)：`output/slides.pptx` (粉彩柔和教育風)
- 網頁版互動式簡報 (HTML)：`output/slides.html` (支援鍵盤左右鍵切換)
- 核心概念社群圖：`output/concept_post.png`

## 使用的輔助腳本與工具

我們建立並部署了一個自動化輔助腳本：`scripts/classroom_analyzer_helper.py`。
它提供以下命令列功能：
- `fetch-audio <URL> --output <mp3_path>`：自 YouTube 影片下載並轉換音軌為 MP3。
- `transcribe <audio_path> --output-srt <srt_path> --output-txt <txt_path>`：使用 Whisper 模型進行高精度語音轉譯。
- `generate-slides --analysis <txt_path> --output <pptx_path> --style <theme>`：將分析報告轉換為 PPTX 簡報與網頁互動式簡報。
- `generate-image --text <txt_path> --output <png_path> --style <theme>`：利用 Pillow 產生核心概念摘要社群圖。

## 執行依賴
- **Python 3.10+** (推薦使用 `uv` 進行執行管理)
- **FFmpeg** (音訊分離與處理)
- **Python 依賴套件**：`python-pptx`, `faster-whisper`, `pillow`, `yt-dlp`
