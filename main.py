from cProfile import label
from tkinter.ttk import Style
from docx2pdf import convert
from numpy import size 
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
        my_btn = wx.Button(panel, label='Upload')
        my_btn.Bind(wx.EVT_BUTTON,self.findDocxFile)
        my_sizer.Add(my_btn, 0, wx.ALL | wx.CENTER, 5)        
        panel.SetSizer(my_sizer)  

        self.Show()



    def findDocxFile(self, event):
        openFileDialog = wx.FileDialog(frame,"Open","","","Docx Files (*.docx)|*.docx", wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        openFileDialog.ShowModal()
        filePath = openFileDialog.GetPath()
        openFileDialog.Destroy()
        progressMax = 100
        docxToPdf(filePath)
        wx.MessageBox('Convert Completed', 'Dialog', wx.OK | wx.ICON_MASK)
        self.Close()

# docx to pdf 
def docxToPdf(docxFile,saveLocation = None):
    print("Starting Convertion")
    convert(docxFile)



# main
if __name__ == '__main__':
    app = wx.App()
    frame = MyFrame()
    app.MainLoop()