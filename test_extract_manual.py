
import os
from extractor import extract_from_file

files = [
    r"c:\Users\muraliba\PycharmProjects\my-project\S4C-Processed-Documents\Hamill5e9781975144654-ch003_KK.docx",
    r"c:\Users\muraliba\PycharmProjects\my-project\S4C-Processed-Documents\Hamill5e9781975144654-ch004_kk.docx"
]

with open("test_output.txt", "w", encoding="utf-8") as out:
    for f in files:
        out.write(f"\nScanning: {os.path.basename(f)}\n")
        try:
            results = extract_from_file(f)
            out.write(f"Found {len(results)} items.\n")
            for item in results:
                out.write(f"[{item['chapter']}] {item['item_type']} {item['item_no']} -> Credit: {item['credit']}\n")
        except Exception as e:
            out.write(f"Error processing {f}: {e}\n")
