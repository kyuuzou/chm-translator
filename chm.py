import os
import chardet
import shutil
import subprocess
import argparse
from bs4 import BeautifulSoup
from bs4.element import Declaration, Doctype
from googletrans import Translator

# Paths to HTML Help Workshop tools
HH_DECOMPILER_PATH = r"C:\Windows\hh.exe"
HH_COMPILER_PATH = r"C:\Program Files (x86)\HTML Help Workshop\hhc.exe"

# Function to decompile CHM file
def decompile_chm(chm_file, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    print(f"Decompiling {chm_file} to {output_dir}...")
    subprocess.run([HH_DECOMPILER_PATH, '-decompile', output_dir, chm_file], check=True)
    print(f"Decompile completed for {chm_file}")

    # Rename .hhc and .hhk files to target.hhc and target.hhk
    for root, _, files in os.walk(output_dir):
        for file in files:
            if file.endswith('.hhc'):
                os.rename(os.path.join(root, file), os.path.join(root, 'target.hhc'))
            elif file.endswith('.hhk'):
                os.rename(os.path.join(root, file), os.path.join(root, 'target.hhk'))

# Function to translate HTML content
def translate_html(content, src='ja', dest='en'):
    print("Translating HTML content...")
    translator = Translator()
    soup = BeautifulSoup(content, 'html.parser')

    print("Finding strings in HTML...")
    for tag in soup.find_all(string=True):
        if (len(tag) < 2):
            continue

        if isinstance(tag, (Declaration, Doctype)):
            continue

        print("Processing string:", tag.string)

        if tag.string and tag.parent.name not in ['script', 'style', 'img']:
            original_text = tag.string
            try:
                translated_text = translator.translate(original_text, src=src, dest=dest).text
            except Exception as e:
                print(f"Translation failed for text '{original_text}'. Error: {e}")
                translated_text = original_text  # Fallback to original text if translation fails
            print(f"Translated text: '{original_text}' -> '{translated_text}'")
            tag.string.replace_with(translated_text)

    print("Translation completed.")
    return str(soup)

# Function to process HTML files in a directory
def process_html_files(input_dir, output_dir, src_lang='ja', dest_lang='en'):
    print("Processing HTML files...")
    for root, dirs, files in os.walk(input_dir):
        for file in files:
            if file.endswith('.html'):
                input_path = os.path.join(root, file)
                output_path = os.path.join(output_dir, os.path.relpath(input_path, input_dir))
                os.makedirs(os.path.dirname(output_path), exist_ok=True)

                try:
                    print(f"Reading file: {input_path}")
                    with open(input_path, 'rb') as f:
                        raw_data = f.read()
                        result = chardet.detect(raw_data)
                        encoding = result['encoding']
                        print(f"Detected encoding: {encoding}")
                        content = raw_data.decode(encoding)
                except Exception as e:
                    print(f"Failed to read {input_path} with detected encoding: {e}")
                    continue

                print(f"Translating file: {input_path}")
                translated_content = translate_html(content, src=src_lang, dest=dest_lang)

                print(f"Writing translated content to: {output_path}")
                with open(output_path, 'w', encoding='utf-8') as f:
                    f.write(translated_content)

    print("HTML files processing completed.")

def translate_hhc_file(file_path, src='ja', dest='en'):
    print(f"Translating .hhc file: {file_path}")
    translator = Translator()

    try:
        with open(file_path, 'rb') as f:
            raw_data = f.read()
            result = chardet.detect(raw_data)
            encoding = result['encoding']
            print(f"Detected encoding: {encoding}")
            content = raw_data.decode(encoding)

        soup = BeautifulSoup(content, 'html.parser')
        for obj in soup.find_all('object'):
            for param in obj.find_all('param', {'name': 'Name'}):
                if param.get('value'):
                    original_text = param['value']
                    try:
                        translated_text = translator.translate(original_text, src=src, dest=dest).text
                        param['value'] = translated_text
                        print(f"Translated text: '{original_text}' -> '{translated_text}'")
                    except Exception as e:
                        print(f"Translation failed for text '{original_text}'. Error: {e}")
                        param['value'] = original_text  # Fallback to original text if translation fails

        translated_content = str(soup)
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(translated_content)

        print(f"Translation completed for .hhc file: {file_path}")

    except Exception as e:
        print(f"Failed to translate .hhc file {file_path}: {e}")

def copy_and_translate_additional_files(decompiled_dir, translated_dir):
    for root, dirs, files in os.walk(decompiled_dir):
        for dir in dirs:
            src_dir_path = os.path.join(root, dir)
            dst_dir_path = os.path.join(translated_dir, os.path.relpath(src_dir_path, decompiled_dir))
            os.makedirs(dst_dir_path, exist_ok=True)
            print(f"Created directory {dst_dir_path}")

        for file in files:
            if not file.endswith('.html'):
                file_path = os.path.join(root, file)
                target_path = os.path.join(translated_dir, os.path.relpath(file_path, decompiled_dir))
                os.makedirs(os.path.dirname(target_path), exist_ok=True)
                shutil.copyfile(file_path, target_path)
                print(f"Copied {file} to {target_path}")

                if file.endswith('.hhc'):
                    translate_hhc_file(target_path, src='ja', dest='en')

def generate_hhp_file(translated_dir, output_file):
    print(f"Generating HHP file in {translated_dir}...")
    hhc_file = None
    hhk_file = None
    html_files = []

    for root, dirs, files in os.walk(translated_dir):
        for file in files:
            if file.endswith('target.hhc'):
                hhc_file = os.path.join(root, file)
            elif file.endswith('target.hhk'):
                hhk_file = os.path.join(root, file)
            elif file.endswith('.html'):
                html_files.append(os.path.relpath(os.path.join(root, file), translated_dir))

    if not hhc_file or not hhk_file:
        raise FileNotFoundError("Missing .hhc or .hhk file in the translated directory.")

    hhp_content = [
        '[OPTIONS]',
        'Compatibility=1.1 or later',
        f'Compiled file={output_file}',
        f'Contents file={os.path.relpath(hhc_file, translated_dir)}',
        f'Index file={os.path.relpath(hhk_file, translated_dir)}',
        '[FILES]'
    ]

    for html_file in html_files:
        hhp_content.append(html_file)

    hhp_path = os.path.join(translated_dir, 'project.hhp')
    with open(hhp_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(hhp_content))

    print(f"HHP file generated at {hhp_path}")
    return hhp_path

def compile_chm(translated_dir, output_chm):
    hhp_file = generate_hhp_file(translated_dir, 'project.hhp')

    # Ensure the hhp file exists
    if not os.path.isfile(hhp_file):
        raise FileNotFoundError(f"HHP file not found: {hhp_file}")

    # Compile the CHM file and capture output and errors
    try:
        result = subprocess.run(
            [HH_COMPILER_PATH, hhp_file],
            check=False,  # Set to False to handle return code manually
            capture_output=True,
            text=True
        )
        print(f"Compilation output:\n{result.stdout}")
        print(f"Compilation errors:\n{result.stderr}")

        # Check the return code and handle errors
        if result.returncode != 0:
            print(f"Warning: hhc.exe returned non-zero exit status {result.returncode}")
            if "error" in result.stderr.lower():
                raise subprocess.CalledProcessError(result.returncode, result.args, output=result.stdout, stderr=result.stderr)

    except subprocess.CalledProcessError as e:
        print(f"Error during compilation:\n{e.stdout}\n{e.stderr}")
        raise

    # Move the compiled CHM to the desired output location
    if os.path.isfile(os.path.join(translated_dir, 'project.hhp')):
        shutil.move(os.path.join(translated_dir, 'project.hhp'), output_chm)
        print(f"Compiled CHM renamed to: {output_chm}")
    else:
        raise FileNotFoundError(f"Compiled CHM not found: {hhp_file}")

def translate_chm(input_chm, temp_dir='temp_chm'):
    decompile_dir = os.path.join(temp_dir, 'decompiled')
    translated_dir = os.path.join(temp_dir, 'translated')

    # Define the temporary and output CHM file names
    temp_chm = os.path.join(temp_dir, 'target.chm')
    file_name = os.path.basename(input_chm)
    output_chm = f"[EN] {file_name}"

    # Clean up any leftovers from interrupted runs
    print(f"Cleaning up temporary directory {temp_dir}...")
    shutil.rmtree(temp_dir, ignore_errors=True)
    os.makedirs(temp_dir, exist_ok=True)

    # Make a temporary copy of the original CHM file and rename it to target.chm
    shutil.copyfile(input_chm, temp_chm)

    # Decompile the CHM file
    decompile_chm(temp_chm, decompile_dir)

    # Translate the HTML files
    process_html_files(decompile_dir, translated_dir)

    # Copy and translate additional files (images, CSS, .hhc, .hhk)
    copy_and_translate_additional_files(decompile_dir, translated_dir)

    # Compile the translated files into a new CHM file
    compile_chm(translated_dir, temp_chm)

    # Rename the temporary CHM file back to the original name
    shutil.move(temp_chm, output_chm)
    print(f"Compiled and translated CHM saved as: {output_chm}")

    # Clean up temporary directories
    print(f"Cleaning up temporary directory {temp_dir}...")
    shutil.rmtree(temp_dir)
    print("Cleanup completed.")

def translate_chm_files_recursively(input_dir):
    for root, dirs, files in os.walk(input_dir):
        for file in files:
            if file.endswith('.chm'):
                chm_file = os.path.join(root, file)

                # Translate the CHM file
                translate_chm(chm_file)

def main(args):
    if args.recursive:
        translate_chm_files_recursively(args.input_chm)
    else:
        translate_chm(args.input_chm)

# Entry point
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Translate and recompile CHM files.")
    parser.add_argument("input_chm", help="Path to the input CHM file")
    parser.add_argument("--recursive", action="store_true", help="Translate CHM files recursively in a directory")
    args = parser.parse_args()

    main(args)
