# Master Design System: Regional Economic Trends Report

## 1. Global Visual Standards (All Pages)
- **Page Specification:**
  - Standard: A4 Portrait (210mm x 297mm).
  - Container: `width: 210mm; min-height: 297mm; padding: 20mm 15mm; margin: 0 auto; background: white; box-sizing: border-box;`
- **Typography:**
  - Font Family: 'Malgun Gothic', 'Dotum', sans-serif.
  - Base Size: 10pt (Body), 11pt (Table Body), 9pt (Dense Table).
  - Line Height: 1.5 ~ 1.6.
- **Color Palette:**
  - Primary (Header): `#2B4E88` (Navy Blue) or `#000000`.
  - Background (Th): `#F5F7FA` (Light Gray/Blue).
  - Accent (Negative): `#E04F4F` (Red text for minus values).
  - Border (Strong): `#000000` (2px Solid).
  - Border (Light): `#DDDDDD` (1px Solid).

## 2. Universal Component: Tables
*Critical Rule: Replicate the "Korean Government Report" table style.*
- **Structure:** `border-collapse: collapse; width: 100%;`
- **Borders:**
  - **Top/Bottom:** `2px solid var(--b-strong)`.
  - **Header Separator:** `1px solid #888` (between th and td).
  - **Inner Grid:** `1px solid var(--b-light)` (check image for dotted vs solid).
- **Alignment:**
  - `th`: Center / Middle.
  - `td` (Text): Center.
  - `td` (Number): Right (`padding-right: 4px`).

## 3. Page Type Specifications (Conditional Styling)

### TYPE A: Summary & Sector Trends (요약, 부문별)
*Files: 실업률.png, 수출.png, 요약-*.png*
- **Layout:** Vertical Stack or 2-Column Grid.
  - Top: Title + Text Summary.
  - Bottom/Right: Chart Placeholder.
- **Charts:** Do NOT render canvas. Use `<div class="chart-placeholder">Chart Area</div>`.
  - Style: `height: 250px; background: #fafafa; border: 1px dashed #ccc; display: flex; justify-content: center; align-items: center; color: #999;`

### TYPE B: Regional Dashboard (시도별 - 서울, 부산 등)
*Files: 서울.png, 참고-GRDP.png*
- **Layout:** Title → Text Summary Box → Key Indicators Table → Chart.
- **Summary Box:** Rounded corners (8px), Background `#f9f9f9`, Padding `15px`.
- **Table Density:** Medium.

### TYPE C: Statistical Tables (통계표)
*Files: 통계표-*.png*
- **Layout:** Full-width dense grid.
- **Font Size:** Reduce to **8pt ~ 9pt** to prevent overflow.
- **Headers:** Complex `rowspan` and `colspan` are mandatory. Replicate exactly.
- **Data:** High density. Ensure numbers are strictly right-aligned.

### TYPE D: Appendix (부록)
*Files: 부록-*.png*
- **Structure:** Definition List (`<dl>`) or Unordered List (`<ul>`).
- **Style:**
  - Term (`dt`): Bold, Bullet point (■ or □).
  - Desc (`dd`): Indented, `padding-left: 1em`.

## 4. Implementation Rules
1. **Source of Truth:** The PNG image dictates spacing and border thickness.
2. **Data Handling:** No dynamic fetching. Hardcode static structure/text.
3. **No Overflow:** Content must fit within the 210mm width.