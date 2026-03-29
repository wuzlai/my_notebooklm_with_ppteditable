import zipfile
import os
import xml.dom.minidom

def inspect_pptx_slide_xml(pptx_path: str, slide_num: int = 1):
    if not os.path.exists(pptx_path):
        print(f"File not found: {pptx_path}")
        return

    try:
        with zipfile.ZipFile(pptx_path, 'r') as zp:
            slide_xml_path = f'ppt/slides/slide{slide_num}.xml'
            if slide_xml_path in zp.namelist():
                content = zp.read(slide_xml_path)
                dom = xml.dom.minidom.parseString(content)
                pretty_xml = dom.toprettyxml()
                print(f"--- XML for {slide_xml_path} ---")
                # Print first 2000 chars to avoid overwhelming
                print(pretty_xml[:2000])
                print("--- ... ---")
                
                # Also check chart XML if exists
                chart_path = f'ppt/charts/chart1.xml'
                if chart_path in zp.namelist():
                    chart_content = zp.read(chart_path)
                    chart_dom = xml.dom.minidom.parseString(chart_content)
                    print(f"\n--- XML for {chart_path} ---")
                    print(chart_dom.toprettyxml()[:2000])
            else:
                print(f"Slide {slide_num} XML not found in ZIP.")
    except Exception as e:
        print(f"Error inspecting PPTX: {e}")

if __name__ == "__main__":
    path = r"projects\测试DEMO\最终文档\测试DEMO_test_hardened_page_01.pptx"
    inspect_pptx_slide_xml(path)
