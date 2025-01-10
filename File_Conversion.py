import os
import subprocess

def convert_file(file_path, input_folder):
    print(f"Converting file: {file_path}")  # 打印正在转换的文件
    if file_path.endswith('.doc'):
        # 如果文件是.doc格式，使用LibreOffice将其转换为.docx格式
        subprocess.run(['/Applications/LibreOffice.app/Contents/MacOS/soffice', '--headless', '--convert-to', 'docx', file_path, '--outdir', input_folder], check=True)
    elif file_path.endswith('.xls'):
        # 如果文件是.xls格式，使用LibreOffice将其转换为.xlsx格式
        subprocess.run(['/Applications/LibreOffice.app/Contents/MacOS/soffice', '--headless', '--convert-to', 'xlsx', file_path, '--outdir', input_folder], check=True)

def convert_files(input_folder):
    for filename in os.listdir(input_folder):
        file_path = os.path.join(input_folder, filename)
        if filename.endswith('.doc') or filename.endswith('.xls'):
            convert_file(file_path, input_folder)  # 转换文件

input_folder = '/Users/hang/Library/Containers/com.tencent.xinWeChat/Data/Library/Application Support/com.tencent.xinWeChat/2.0b4.0.9/237e6ba0b2b7af52a974319f16a23be7/Message/MessageTemp/c1c30910da9401121316d3c0d92ec2ed/File'
convert_files(input_folder)  # 转换指定文件夹中的所有文件