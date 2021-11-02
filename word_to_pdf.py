from docx2pdf import convert
import PyPDF2
import os
import glob

path = os.getcwd()
output_dir = path+'\converted_files'
if not os.path.isdir(output_dir):
  os.mkdir(output_dir)

for dir in glob.glob(path+'\**', recursive=True):
  if r'.' in dir: #'.'が含まれていたら, ファイルと判定する
    continue
  convert(dir, output_dir) # convert(input_dir, output_dir)

merger = PyPDF2.PdfFileMerger()
for file in glob.glob(output_dir + '\*.pdf'):
  if 'combined' in file:
    continue
  merger.append(file)
merger.write(path + '\combined.pdf')
merger.close()
