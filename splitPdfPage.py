import PyPDF2
from tkinter import *
from tkinter.filedialog import askopenfilename


class MainWindow:
    def __init__(self):
        self.tk = Tk()
        self.tk.title("PDF操作器 (dev by: Hanqi)")
        self.font = ('微软雅黑', 12)

        self.frame_btn()
        self.frame_split_pdf()
        self.frame_split_page()

        self.text_frame = Frame(self.tk)
        self.text = Text(self.text_frame, font=self.font, height=10)
        self.text.configure(state=DISABLED)
        self.text_frame.pack(side=TOP, fill=BOTH, expand=YES, padx=5, pady=5)
        self.text.pack(side=LEFT, fill=BOTH, expand=YES, ipadx=2, ipady=2)

        self.tk.mainloop()

    def frame_btn(self):
        self.bnt_frame = Frame(self.tk)
        self.input_btn = Button(self.bnt_frame, text="选择PDF文件", font=self.font, command=self.input_btn_command)
        self.bnt_frame.pack(side=TOP, fill=X, expand=YES, padx=5, pady=5)
        self.input_btn.pack(side=LEFT, fill=X, expand=YES, ipadx=2, ipady=2, padx=2, pady=2)

    def frame_split_pdf(self):
        self.split_frame = Frame(self.tk)
        self.frame_label = Label(self.split_frame, text="分割PDF：", font=self.font)
        self.from_label = Label(self.split_frame, text="起始页: ", font=self.font)
        self.to_label = Label(self.split_frame, text="最终页（包括）: ", font=self.font)
        self.from_entry = Entry(self.split_frame, font=self.font)
        self.to_entry = Entry(self.split_frame, font=self.font)
        self.start_btn = Button(self.split_frame, text="开始", font=self.font, command=self.start_split_pdf)
        self.split_frame.pack(side=TOP, fill=X, expand=YES, padx=5, pady=5)
        self.frame_label.pack(side=LEFT, ipadx=2, ipady=2, padx=2, pady=2)
        self.from_label.pack(side=LEFT, ipadx=2, ipady=2, padx=2, pady=2)
        self.from_entry.pack(side=LEFT, ipadx=2, ipady=2, padx=2, pady=2)
        self.to_label.pack(side=LEFT, ipadx=2, ipady=2, padx=2, pady=2)
        self.to_entry.pack(side=LEFT, ipadx=2, ipady=2, padx=2, pady=2)
        self.start_btn.pack(side=LEFT, fill=X, expand=YES, ipadx=2, ipady=2, padx=2, pady=2)

    def frame_split_page(self):
        self.split_frame = Frame(self.tk)
        self.frame_label = Label(self.split_frame, text="分割PDF首页: ", font=self.font)
        self.top_label = Label(self.split_frame, text="上: ", font=self.font)
        self.bottom_label = Label(self.split_frame, text="下: ", font=self.font)
        self.left_label = Label(self.split_frame, text="左: ", font=self.font)
        self.right_label = Label(self.split_frame, text="右: ", font=self.font)
        self.top_entry = Entry(self.split_frame, font=self.font)
        self.bottom_entry = Entry(self.split_frame, font=self.font)
        self.left_entry = Entry(self.split_frame, font=self.font)
        self.right_entry = Entry(self.split_frame, font=self.font)
        self.start_btn = Button(self.split_frame, text="开始", font=self.font, command=self.start_split_page)
        self.split_frame.pack(side=TOP, fill=X, expand=YES, padx=5, pady=5)
        self.frame_label.pack(side=LEFT, ipadx=2, ipady=2, padx=2, pady=2)
        self.top_label.pack(side=LEFT, ipadx=2, ipady=2, padx=2, pady=2)
        self.top_entry.pack(side=LEFT, ipadx=2, ipady=2, padx=2, pady=2)
        self.bottom_label.pack(side=LEFT, ipadx=2, ipady=2, padx=2, pady=2)
        self.bottom_entry.pack(side=LEFT, ipadx=2, ipady=2, padx=2, pady=2)
        self.left_label.pack(side=LEFT, ipadx=2, ipady=2, padx=2, pady=2)
        self.left_entry.pack(side=LEFT, ipadx=2, ipady=2, padx=2, pady=2)
        self.right_label.pack(side=LEFT, ipadx=2, ipady=2, padx=2, pady=2)
        self.right_entry.pack(side=LEFT, ipadx=2, ipady=2, padx=2, pady=2)
        self.start_btn.pack(side=LEFT, fill=X, expand=YES, ipadx=2, ipady=2, padx=2, pady=2)

    def input_btn_command(self):
        self.pdf_path = askopenfilename(title="选择PDF文件", filetypes=(("选择PDF文件", "*.pdf"),))
        self.text_print(self.pdf_path)

    def start_split_pdf(self):
        input_file = PyPDF2.PdfFileReader(self.pdf_path)
        output_file = PyPDF2.PdfFileWriter()

        start = int(self.from_entry.get())
        end = int(self.to_entry.get())

        for page in range(start-1, end):
            output_file.addPage(input_file.getPage(page))
        output_file.write(open("./output.pdf", 'wb'))
        self.text_print('DONE!')

    def start_split_page(self):
        input_file = PyPDF2.PdfFileReader(open(self.pdf_path, 'rb'))
        output_file = PyPDF2.PdfFileWriter()

        page = input_file.getPage(0)
        width = float(page.mediaBox.getWidth())
        height = float(page.mediaBox.getHeight())

        left_margin = int(self.left_entry.get())
        right_margin = int(self.right_entry.get())
        top_margin = int(self.top_entry.get())
        bottom_margin = int(self.bottom_entry.get())

        page.mediaBox.lowerLeft = (left_margin, bottom_margin)
        page.mediaBox.lowerRight = (width - right_margin, bottom_margin)
        page.mediaBox.upperLeft = (left_margin, height - top_margin)
        page.mediaBox.upperRight = (width - right_margin, height - top_margin)

        output_file.addPage(page)
        output_file.write(open("./output.pdf", 'wb'))
        self.text_print('DONE!')

    def text_print(self, str):
        self.text.configure(state=NORMAL)
        self.text.insert(INSERT, str + "\n")
        self.text.configure(state=DISABLED)
        self.text.see(END)


if __name__ == '__main__':
    MainWindow()
