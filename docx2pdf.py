import os
import subprocess
import logging
from tqdm import tqdm

def convert_docx_to_pdf(input_folder):
    # ロギングの設定
    logging.basicConfig(filename='conversion_errors.log', level=logging.ERROR)
    
    # inputフォルダ内の.docxファイルのリストを取得
    docx_files = [f for f in os.listdir(input_folder) if f.endswith(".docx")]

    # LibreOfficeのフルパスを設定（whichコマンドで確認したパスを使用）
    libreoffice_path = '/opt/homebrew/bin/soffice'
    
    # プログレスバー
    with tqdm(total=len(docx_files), desc="DOCXをPDFに変換中") as pbar:
        for filename in docx_files:
            try:
                doc_path = os.path.join(input_folder, filename)
                pdf_path = os.path.join(input_folder, f"{os.path.splitext(filename)[0]}.pdf")
                
                # LibreOfficeを使用してDOCXをPDFに変換
                result = subprocess.run(
                    [libreoffice_path, '--headless', '--convert-to', 'pdf', '--outdir', input_folder, doc_path],
                    capture_output=True, text=True
                )

                if result.returncode != 0:
                    logging.error(f"{filename}の変換に失敗しました: {result.stderr}")
                
            except Exception as e:
                logging.error(f"{filename}の処理中にエラーが発生しました: {str(e)}")
            
            pbar.update(1)

input_folder = 'input'  # Wordファイルを含むフォルダを指定
convert_docx_to_pdf(input_folder)