import os
import subprocess

# 定义源文件夹和目标文件夹路径
source_folder = '/Users/hang/Library/Containers/com.tencent.xinWeChat/Data/Library/Application Support/com.tencent.xinWeChat/2.0b4.0.9/237e6ba0b2b7af52a974319f16a23be7/Message/MessageTemp/c1c30910da9401121316d3c0d92ec2ed/File'  # 替换为你的源文件夹路径
destination_folder = '/Users/hang/Library/Containers/com.tencent.xinWeChat/Data/Library/Application Support/com.tencent.xinWeChat/2.0b4.0.9/237e6ba0b2b7af52a974319f16a23be7/Message/MessageTemp/c1c30910da9401121316d3c0d92ec2ed/File'  # 替换为你的目标文件夹路径

# 遍历源文件夹中的所有文件
for root, folders, files in os.walk(source_folder):
    for file in files:
        if file.endswith(".doc"):
            file_path = os.path.abspath(os.path.join(root, file))
            output_path = os.path.join(destination_folder, os.path.splitext(file)[0] + ".docx")
            subprocess.run(["/Applications/LibreOffice.app/Contents/MacOS/soffice", "--headless", "--convert-to", "docx", file_path, "--outdir", os.path.dirname(output_path)])
            # 如果需要删除源文件，可以取消以下行的注释
            # os.remove(file_path)
        elif file.endswith(".xls"):
            file_path = os.path.abspath(os.path.join(root, file))
            output_path = os.path.join(destination_folder, os.path.splitext(file)[0] + ".xlsx")
            subprocess.run(["/Applications/LibreOffice.app/Contents/MacOS/soffice", "--headless", "--convert-to", "xlsx", file_path, "--outdir", os.path.dirname(output_path)])
            # 如果需要删除源文件，可以取消以下行的注释
            # os.remove(file_path)

print('Conversion completed!')