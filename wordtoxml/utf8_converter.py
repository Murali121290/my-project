import sys
import chardet
import os
import re

def convert_to_utf8(file_path):
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return

    # Read the file content
    with open(file_path, 'rb') as f:
        raw_data = f.read()

    # Detect encoding (handles UTF-16LE, UTF-16BE, UTF-8, Windows-1252, etc.)
    result = chardet.detect(raw_data)
    encoding = result['encoding']
    
    if encoding is None:
        encoding = 'utf-8' # Fallback
    
    print(f"Detected encoding for {file_path}: {encoding}")

    try:
        # Decode the content
        content = raw_data.decode(encoding, errors='replace')
        
        # If XML/HTML file, ensure unescaped ampersands and non-ASCII special characters (smart quotes, dashes, etc.) are converted to XML entities &#xXXXX;
        if file_path.lower().endswith(('.xml', '.postxml', '.posthtml', '.html')):
            content = re.sub(r'&(?!amp;|lt;|gt;|quot;|apos;|#[0-9]+;|#x[0-9a-fA-F]+;)', '&#x0026;', content)
            content = re.sub(r'[^\x00-\x7F]', lambda m: f"&#x{ord(m.group(0)):04X};", content)
        
        # Write back as UTF-8 (without BOM)
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"Successfully converted {file_path} to UTF-8 with entity mapping")
    except Exception as e:
        print(f"Error converting {file_path}: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python utf8_converter.py <file_path>")
    else:
        convert_to_utf8(sys.argv[1])


