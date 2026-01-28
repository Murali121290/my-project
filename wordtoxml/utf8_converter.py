import sys
import chardet
import os

def convert_to_utf8(file_path):
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return

    # Read the file content
    with open(file_path, 'rb') as f:
        raw_data = f.read()

    # Detect encoding
    result = chardet.detect(raw_data)
    encoding = result['encoding']
    
    if encoding is None:
        encoding = 'utf-8' # Fallback
    
    print(f"Detected encoding for {file_path}: {encoding}")

    try:
        # Decode the content
        content = raw_data.decode(encoding)
        
        # Write back as UTF-8 (without BOM)
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"Successfully converted {file_path} to UTF-8")
    except Exception as e:
        print(f"Error converting {file_path}: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python utf8_converter.py <file_path>")
    else:
        convert_to_utf8(sys.argv[1])
