from os import listdir
from os.path import isfile, join
import comtypes.client
from tqdm import tqdm

class ConvertToPdf():
    def __init__(self, input_folder):
        self.input_folder = input_folder
        self.word = comtypes.client.CreateObject("Word.Application")
        self.powerpoint = comtypes.client.CreateObject('PowerPoint.Application')
        self.excel = comtypes.client.CreateObject('Excel.Application')
        self.files = []
        self.logs = []

    def convert(self):
        try:
            files = self.get_files()
            print(files)
            for file in tqdm(files, desc='Converting files: ', unit="files"):
                full_path = self.input_folder + '\\' + file
                if file.endswith('.doc') or file.endswith('.docx'):
                    self.convert_doc_to_pdf(full_path)
                elif file.endswith('.ppt') or file.endswith('.pptx'):
                    self.convert_ppt_to_pdf(full_path)
                elif file.endswith('.xls') or file.endswith('.xlsx'):
                    self.convert_excel_to_pdf(full_path)
                else:
                    continue
            if self.logs:
                self.save_logs()
        except Exception as e:
            print(f'Error in converting files: {e}')

    def get_files(self):
        try:
            files = [f for f in listdir(self.input_folder) if isfile(join(self.input_folder, f))]
            self.files = files
            return files
        except Exception as e:
            self.logs.append(f'Error in getting files: {e}')
            return None
    
    def convert_ppt_to_pdf(self, file):
        try:
            powerpoint = self.powerpoint
            pdf = powerpoint.Presentations.Open(file)
            powerpoint.DisplayAlerts = 0
            pdf.SaveAs(file.replace(self.get_fileformat(file), '.pdf'), 32)
            pdf.Close()
            powerpoint.Quit()
        except Exception as e:
            self.logs.append(f'Error in converting {file} to PDF: {e}')
            print(f'Error PowerPoint file in converting {file} to PDF: {e}')

    def convert_doc_to_pdf(self, file):
        try:
            word = self.word
            pdf = word.Documents.Open(file)
            word.DisplayAlerts = 0
            pdf.SaveAs(file.replace(self.get_fileformat(file), '.pdf'), 17)
            pdf.Close()
            word.Quit()
        except Exception as e:
            self.logs.append(f'Error in converting {file} to PDF: {e}')
            print(f'Error Word file in converting {file} to PDF: {e}')

    def convert_excel_to_pdf(self, file):
        try:
            excel = self.excel
            pdf = excel.Workbooks.Open(file)
            excel.DisplayAlerts = 0 
            pdf.SaveAs(file.replace(self.get_fileformat(file), '.pdf'), 57)
            pdf.Close()
            excel.Quit()
        except Exception as e:
            self.logs.append(f'Error in converting {file} to PDF: {e}')
            print(f'Error Excel file in converting {file} to PDF: {e}')

    def get_fileformat(self, file):
        return '.' + file.split('.')[-1].lower()
    
    def save_logs(self):
        with open(self.input_folder + '\\ + ''logs.txt', 'w') as f:
            for log in self.logs:
                f.write(log + '\n')
    
if __name__ == '__main__':
    input_folder = input('Enter the input folder path: ')
    convert_to_pdf = ConvertToPdf(input_folder)
    convert_to_pdf.convert()
        