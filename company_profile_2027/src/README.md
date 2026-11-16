# 회사소개서 생성 스크립트

`company_profile_2027/` 폴더의 PPT 2종과 엑셀 양식을 만드는 원본입니다.

## 다시 만들기

```bash
npm install
node build.js
python3 xlsx_build.py ../TOIN_회사소개서_자료수집양식.xlsx
```

- `node build.js` → 이 폴더에 `TOIN_회사소개서_2027_KR.pptx`, `TOIN_Company_Profile_2027_EN.pptx` 생성. 확인 후 상위 폴더로 복사.
- 엑셀은 `openpyxl` 필요 (`pip install openpyxl`).
- 이미지는 `img/`, 이미지 크기 정보는 `dims.json` (새 이미지를 넣으면 크기 추가 필요).
- 슬라이드 내용·발표자 노트는 `build.js`, 엑셀 시트 구성은 `xlsx_build.py`에서 수정.
