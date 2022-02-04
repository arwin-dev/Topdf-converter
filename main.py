from docx2pdf import convert 
import wx

# GUI
class MyFrame(wx.Frame):    
    def __init__(self):
        super().__init__(parent=None, title='Hello World')
        self.Show()

# docx to pdf 
def docxToPdf(file,save):
    print("Starting Convertion")
    convert("C:/Users/Arwin/Documents/_projects/Topdf-converter/123.docx","C:/Users/Arwin/Desktop/comple.pdf")

# main
if __name__ == '__main__':
    app = wx.App()
    frame = MyFrame()
    app.MainLoop()