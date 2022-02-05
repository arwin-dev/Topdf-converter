from docx2pdf import convert
from pdf2docx import converter
import win32com.client
import wx

# GUI
class MyFrame(wx.Frame):    
    def __init__(self):
        super().__init__(parent=None, title='Docx to PDF converter')
        panel = wx.Panel(self)
        my_sizer = wx.BoxSizer(wx.VERTICAL) 

        font = wx.Font(20, family = wx.FONTFAMILY_ROMAN, style = 0, weight = 90,underline = False, faceName ="", encoding = wx.FONTENCODING_DEFAULT)

        mesg = wx.StaticText(panel,label="Select file to convert", pos = (100,50), size =wx.DefaultSize, )
        mesg.SetFont(font)
        my_sizer.Add(mesg , wx.ALL | wx.EXPAND, 5)   
        my_btn = wx.Button(panel, label='Convert Docx to PDF')
        my_btn2 = wx.Button(panel, label='Convert PDF to Docx')
        my_btn.Bind(wx.EVT_BUTTON,self.findDocxFile)
        my_btn2.Bind((wx.EVT_BUTTON),self.findPdfFile)
        my_sizer.Add(my_btn, 0, wx.ALL | wx.CENTER, 5)  
        my_sizer.Add(my_btn2,0,wx.ALL | wx.ALIGN_CENTER_HORIZONTAL,5)      
        panel.SetSizer(my_sizer)  

        self.Show()



    def findDocxFile(self, event):
        openFileDialog = wx.FileDialog(frame,"Open","","","Docx Files (*.docx)|*.docx", wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        openFileDialog.ShowModal()
        filePath = openFileDialog.GetPath()
        openFileDialog.Destroy()
        if filePath:
            docxToPdf(filePath)
            wx.MessageBox('Convert Completed', 'Dialog', wx.OK | wx.ICON_MASK)
        else:
            wx.MessageBox('File not selected', 'Dialog', wx.OK | wx.ICON_ERROR)
        self.Close()

    def findPdfFile(self,event):
        openFileDialog = wx.FileDialog(frame,"Open","","","PDF Files (*.pdf)|*.pdf", wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        openFileDialog.ShowModal()
        filePath = openFileDialog.GetPath()
        openFileDialog.Destroy()
        if filePath:
            print("FIle convert here ")
            pdfToDocs(filePath)
            wx.MessageBox('Convert Completed', 'Dialog', wx.OK | wx.ICON_MASK)
        else:
            wx.MessageBox('File not selected', 'Dialog', wx.OK | wx.ICON_ERROR)
        self.Close()    

# docx to pdf 
def docxToPdf(docxFile,saveLocation = None):
    print("Starting Convertion")
    convert(docxFile)

def pdfToDocs(pdfFile,saveLocation = None):
    print("Starting Convertion")
    cv = converter(pdfFile)
    save = replace(".pdf","docx")
    cv.conver(save)
    cv.close()



# main
if __name__ == '__main__':
    app = wx.App()
    frame = MyFrame()
    app.MainLoop()