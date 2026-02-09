# Cyclops

Screen OCR 기반 자동 클릭 매크로. 지정한 화면 영역을 반복 감시하며, 원하는 텍스트가 나타나면 자동으로 멈춥니다.

> macOS / Windows 크로스 플랫폼 지원

---

## How it works

```
┌─────────────────────────────────────────────────┐
│  1. 감시 영역 드래그 선택                            │
│  2. 클릭할 이미지 드래그 캡처                         │
│  3. 매칭할 텍스트 입력                              │
│  4. 클릭 딜레이 설정                                │
│              ▼                                  │
│      [ START ] 클릭                              │
│              ▼                                  │
│   ┌─ OCR 스캔 ─── 텍스트 매칭? ── Yes ──→ STOP      │
│   │                    │                        │
│   │                   No                        │
│   │                    ▼                        │
│   │         이미지 찾기 → 클릭                      │
│   │                    ▼                        │
│   │              딜레이 대기                       │
│   │                    │                        │
│   └────────────────────┘                        │
└─────────────────────────────────────────────────┘
```

## Requirements

### Python 3.8+

### System Dependencies

| | macOS | Windows |
|---|---|---|
| **OCR 엔진** | `brew install tesseract tesseract-lang` | [tesseract 다운로드](https://github.com/UB-Mannheim/tesseract/wiki) |

### Python Packages

```bash
pip install pyautogui pytesseract Pillow opencv-python-headless
```

## Usage

```bash
python main.py
```

### Step 1 — 감시 영역 설정
**Set Region** 클릭 → 화면에서 OCR로 감시할 영역을 드래그로 선택합니다.

### Step 2 — 클릭 이미지 캡처
**Capture Image** 클릭 → 반복 클릭할 버튼/아이콘을 드래그로 캡처합니다.

### Step 3 — 매칭 텍스트 입력
**Input Text** 클릭 → OCR에서 감지되면 매크로를 멈출 텍스트를 입력합니다.

### Step 4 — 클릭 딜레이
`+` / `-` 버튼으로 클릭 간격을 조절합니다. (0.5s ~ 10.0s)

### 실행 / 정지
- **START** 버튼 또는 매크로 시작
- **STOP** 버튼 또는 `ESC` 키로 정지
- 마우스를 화면 모서리로 이동하면 **긴급 정지** (PyAutoGUI FAILSAFE)

## Notes

- 이미지 검색은 Step 1에서 지정한 영역 내부에서만 수행됩니다.
- 이미지를 10회 연속 찾지 못하면 자동 종료됩니다.
- **Preview** 버튼으로 캡처된 영역/이미지를 언제든 확인할 수 있습니다.

## License

MIT
