
import sys
from pathlib import Path
import os
import time

# Add project root to path
project_root = Path(__file__).resolve().parent
sys.path.append(str(project_root))

from config.reports import SECTOR_REPORTS
from services.report_generator import generate_report_html
from services.excel_cache import get_excel_file
from generate_full_report import _build_final_html

def extract_year_quarter_from_excel(excel_path_str: str):
    """
    Extracts year and quarter from the Excel filename or defaulting to current config.
    Simple parser for '분석표_25년 3분기...'
    """
    try:
        import re
        basename = os.path.basename(excel_path_str)
        # Match '25년 3분기' or similar
        match = re.search(r'(\d+)년\s*(\d+)분기', basename)
        if match:
            year_short = int(match.group(1))
            quarter = int(match.group(2))
            year = 2000 + year_short
            return year, quarter
    except Exception:
        pass
    return 2025, 3

def main():
    excel_path = "분석표_25년 3분기_캡스톤(업데이트).xlsx"
    if not Path(excel_path).exists():
        print(f"Error: {excel_path} not found")
        return

    print(f"Loading {excel_path}...")
    try:
        excel_file = get_excel_file(excel_path)
    except Exception as e:
        print(f"Error opening Excel file: {e}")
        return

    year, quarter = extract_year_quarter_from_excel(excel_path)
    print(f"Report Period: {year} Q{quarter}")
    
    generated_pages = []
    
    # 1. Generate pages for each sector report
    for config in SECTOR_REPORTS:
        report_id = config['id']
        print(f"\n--- Generating Report: {config['name']} ({report_id}) ---")
        
        try:
            # Generate HTML using service
            html_content, error, missing = generate_report_html(
                excel_path, 
                config, 
                year, 
                quarter, 
                excel_file=excel_file
            )
            
            if html_content:
                print(f"✅ Generated {len(html_content)} bytes")
                generated_pages.append({
                    'title': config['name'],
                    'report_id': report_id,
                    'html': html_content
                })
            elif error:
                print(f"❌ Error generating {report_id}: {error}")
            else:
                print(f"⚠️ No content generated for {report_id}")
                
        except Exception as e:
            print(f"❌ Error generating {report_id}: {e}")
            import traceback
            traceback.print_exc()

    # 2. Build final HTML
    if generated_pages:
        print(f"\nCombining {len(generated_pages)} pages into final report...")
        final_html = _build_final_html(generated_pages, year, quarter)
        
        output_dir = Path("exports")
        output_dir.mkdir(exist_ok=True)
        output_path = output_dir / "test_sector_unified_report.html"
        
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(final_html)
            
        print(f"\n🎉 Success! Report saved to: {output_path}")
        print(f"Total pages: {len(generated_pages)}")
    else:
        print("\n⚠️ Not enough pages generated to build final report.")

if __name__ == "__main__":
    main()
