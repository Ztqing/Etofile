from CustomWindow import CustomWindow
from tkinter import filedialog, messagebox
from ttkbootstrap import Window, Button, StringVar, Entry, Frame, Label, Checkbutton, Labelframe, BooleanVar, \
    Notebook, IntVar, Separator, Text, Scrollbar, Canvas
from ttkbootstrap.constants import LEFT, TOP, RIGHT, BOTTOM, HORIZONTAL, Y, INSERT
import utils


class MainWindow(CustomWindow):
    """主窗口模块"""

    def __init__(self, version, update_time, **kwargs) -> None:

        self.version = version
        self.update_time = update_time

        # 创建窗体
        self.root = Window()

        # 窗体初始化
        super().__init__(self.root, **kwargs)

        """----------------------------------------------布局初始化---------------------------------------------------"""

        # 变量初始化
        self.excel: utils.Excel2Img
        self.flag_date = BooleanVar()  # 以日期为父目录
        self.src_file = StringVar()  # 源文件地址
        self.dst_dir = StringVar()  # 目标文件夹
        self.name_rule = StringVar()  # 命名规则
        self.preview_name = StringVar()  # 预览样式
        self.sheet_id = IntVar()  # 表单编号
        self.sheet_id.set(1)
        self.start_point = StringVar()  # 起点坐标
        self.count = IntVar()  # 链接数量

        # 界面布局
        self.frame = Frame(self.root, padding=10)
        self.frame.pack(side=BOTTOM, fill="both", expand=True)
        self.notebook = Notebook(self.frame, style="light")
        self.notebook.pack(fill="both", expand=True)

        self.main_tab = self.create_main_tab
        self.notebook.add(self.main_tab, text="主页")

        self.about_tab = self.create_about_tab
        self.notebook.add(self.about_tab, text="关于")

        """---------------------------------------------------------------------------------------------------------"""

        self.root.deiconify()

        self.root.mainloop()

    """----------------------------------------------------主页页面---------------------------------------------------"""

    @property
    def create_main_tab(self) -> Frame:
        """
        主页页面布局
        """
        tab = Frame(self.notebook, padding=10)

        """---------------------------------------------------左布局--------------------------------------------------"""
        frame_left = Frame(tab, padding=5)  # 左
        frame_left.pack(side=LEFT, fill="both", expand=True)

        """--------------------------------------------路径设置-------------------------------------------"""

        # 路径（选择文件、文件夹）
        frame_path = Frame(frame_left, padding=5)
        frame_path.pack(side=BOTTOM, expand=True, fill="both")

        # 导出文件路径
        frame_file = Frame(frame_path, padding=5)
        frame_file.pack(expand=True, fill="both")
        label_file = Label(frame_file, style="dark", text="导出文件", padding=5)
        label_file.pack(side=TOP, anchor='w')
        entry_file = Entry(frame_file, textvariable=self.src_file, state="disabled")
        entry_file.pack(side=LEFT, expand=True, padx=5)
        button_file = Button(frame_file, text='更改', command=self.change_file, style="primary-outline")
        button_file.pack(side=LEFT, expand=True, padx=5)

        # 导入目录路径
        frame_dir = Frame(frame_path, padding=5)
        frame_dir.pack(expand=True, fill="both")
        label_dir = Label(frame_dir, style="dark", text="导入目录", padding=5)
        label_dir.pack(side=TOP, anchor='w')
        entry_dir = Entry(frame_dir, textvariable=self.dst_dir, state="disabled")
        entry_dir.pack(side=LEFT, expand=True, padx=5)
        button_dir = Button(frame_dir, text='更改', command=self.change_dir, style="primary-outline")
        button_dir.pack(side=LEFT, expand=True, padx=5)

        """--------------------------------------------命名规则-------------------------------------------"""

        labelframe_name = Labelframe(frame_left, text="文件命名", style="default", padding=5)
        labelframe_name.pack(side=TOP, expand=True, fill="both")

        frame_file_name = Frame(labelframe_name, padding=5)
        frame_file_name.pack(side=TOP, expand=True, fill="both")
        label_name = Label(frame_file_name, text="规范：")
        label_name.pack(side=LEFT, expand=True, padx=5)
        entry_name = Entry(frame_file_name, textvariable=self.name_rule)
        entry_name.focus_set()
        entry_name.bind("<Key>", self.preview)
        # entry_name.bind("<KeyRelease>", self.preview)
        entry_name.pack(side=LEFT, expand=True, padx=5)

        frame_preview = Frame(labelframe_name, padding=5)
        frame_preview.pack(side=TOP, expand=True, fill="both")
        label_preview = Label(frame_preview, text="预览：")
        label_preview.pack(side=LEFT, expand=True, padx=5)
        entry_preview = Entry(frame_preview, textvariable=self.preview_name, state="disabled")
        entry_preview.bind()
        entry_preview.pack(side=LEFT, expand=True, padx=5)

        """---------------------------------------------------右布局--------------------------------------------------"""

        frame_right = Frame(tab, padding=5)  # 右
        frame_right.pack(side=RIGHT, fill="both", expand=True)

        """---------------------------------------------选项--------------------------------------------"""

        labelframe_chose = Labelframe(frame_right, text="选项", style="default", padding=5)
        labelframe_chose.pack(side=TOP, expand=True, fill="both")

        frame_sheet = Frame(labelframe_chose, padding=5)
        frame_sheet.pack(side=TOP, expand=True, fill="both")
        label_sheet = Label(frame_sheet, text="表单编号：")
        label_sheet.pack(side=LEFT, expand=True, padx=5)
        entry_sheet = Entry(frame_sheet, textvariable=self.sheet_id, width=7)
        entry_sheet.bind("<KeyRelease>", self.change_sheet)
        entry_sheet.pack(side=LEFT, expand=True)

        frame_point = Frame(labelframe_chose, padding=5)
        frame_point.pack(side=TOP, expand=True, fill="both")
        label_point = Label(frame_point, text="起点：")
        label_point.pack(side=LEFT, expand=True, padx=5)
        entry_point = Entry(frame_point, textvariable=self.start_point, width=10)
        entry_point.pack(side=LEFT, expand=True)

        frame_count = Frame(labelframe_chose, padding=5)
        frame_count.pack(side=TOP, expand=True, fill="both")
        label_count = Label(frame_count, text="数量：")
        label_count.pack(side=LEFT, expand=True, padx=5)
        entry_count = Entry(frame_count, textvariable=self.count, width=10)
        entry_count.pack(side=LEFT, expand=True)

        frame_date = Frame(labelframe_chose, padding=5)
        frame_date.pack(side=BOTTOM, expand=True, fill="both")
        checkbutton_date = Checkbutton(frame_date, variable=self.flag_date, style="primary-round-toggle",
                                       text="以日期为父目录")
        checkbutton_date.pack(side=LEFT, anchor='w', expand=True, padx=5)

        """---------------------------------------------导出--------------------------------------------"""

        frame_export = Frame(frame_right, padding=5)
        frame_export.pack(side=BOTTOM, expand=True, fill="both")
        # 导出按钮
        button_export = Button(frame_export, text='进行导出', command=self.upload, style="success", width=13)
        button_export.pack(expand=True)

        return tab

    """----------------------------------------------------关于页面---------------------------------------------------"""

    @property
    def create_about_tab(self) -> Frame:
        """
        关于页面布局
        """
        tab = Frame(self.notebook, padding=10)

        """---------------------------------------------------关于---------------------------------------------------"""

        frame_about = Frame(tab, padding=5)
        frame_about.pack(fill="both")

        frame_title = Frame(frame_about)
        frame_title.pack(fill="both")

        title = Label(frame_title, text='Etofile')
        title.config(font=("", 20, "bold"))
        title.pack(side=LEFT, fill="both", pady=5)

        Label(frame_title, text='基于Tkinter开发的Excel超链接文件提取工具。').pack(side=BOTTOM, fill="both", padx=10)

        Separator(frame_about, orient=HORIZONTAL).pack(fill="both", pady=5)

        Label(frame_about, text="版本  {} ({})\t美化  ttkbootstrap".format(self.version, self.update_time)).pack(
            side=LEFT,
            fill="both",
            pady=5)

        """---------------------------------------------------帮助---------------------------------------------------"""

        frame_help = Frame(tab, padding=5)
        frame_help.pack(fill="both")

        text = Text(frame_help, width=0, height=5, spacing3=10, spacing2=10, spacing1=10)
        scroll_help = Scrollbar(frame_help, style="secondary-round")

        text.tag_add('title', INSERT)
        text.tag_config('title', font=("黑体", 10, "bold"), spacing3=0)
        text.tag_add('main', INSERT)
        text.tag_config('main', font=("黑体", 10, ""))
        text.tag_add('link', INSERT)
        text.tag_config('link', font=("黑体", 10, ""))

        str_help = "帮助\n"
        text.insert(INSERT, str_help, 'title')

        can1 = Canvas(text, width=523, height=5)
        can1.create_line(0, 2, 523, 2)

        text.window_create(INSERT, window=can1)

        str_help = "\n该软件是作者使用腾讯收集表过程中，因图片导出重命名需重复作业感到烦恼而制作。\n" \
                   "软件主页由文件命名、导出导入、选项及导出按钮四部分组成：\n" \
                   "命名规范，检测 一对美元符 所包含的单元格（如 c2、D2 等）的内容并与其他进行组合，\n" \
                   "如 xx$c2$x$d3$x ：假设单元格c2、d3处内容分别为“你好”、“世界”，则命名即是 xx你好x世界x ；" \
                   "第2个文件规范则为 xx$c3$x$d4$x 以此类推。\n" \
                   "若命名规范内容为“Dir/xxx”，则将优先创建“Dir”文件夹用于承载“xxx”文件。\n" \
                   "导出导入部分分别指明需要导出的excel类型文件（目前仅支持xlsx类型）以及导入的目录。\n" \
                   "选项部分是用来调节参数，如表单号、起点等，表单号默认为1（即sheet1）；\n" \
                   "起点应为单元格名称（如：e2、D3等）；数量若为12，即由起点向下延伸12个单元格内的超链接文件（包括起点）。\n" \
                   "以日期为父目录选项是针对如收取截图等，需要以日期为目录时省去建立花费。\n"
        text.insert(INSERT, str_help, 'main')

        str_author = "制作\n"
        text.insert(INSERT, str_author, 'title')

        can2 = Canvas(text, width=523, height=5)
        can2.create_line(0, 2, 523, 2)

        text.window_create(INSERT, window=can2)

        str_author = "作者：Tin\n" \
                     "项目地址：https://github.com/Ztqing/Etofile"
        text.insert(INSERT, str_author, 'main')

        # 两个控件关联
        scroll_help.config(command=text.yview)
        text.config(yscrollcommand=scroll_help.set, state="disabled")

        scroll_help.pack(side=RIGHT, fill=Y)
        text.pack(fill="both")

        return tab

    """----------------------------------------------------调用方法---------------------------------------------------"""

    def preview(self, event=None):
        """
        预览文件名
        """
        name_rule = self.name_rule.get()
        try:
            self.preview_name.set(self.excel.get_preview(name_rule))
        except Exception as err:
            print(err)
            self.preview_name.set(name_rule)
        return True

    def upload(self):
        """
        下载文件
        """
        try:
            name_rule = self.name_rule.get()
            file_path = self.src_file.get()
            dir_path = self.dst_dir.get()
            sheet_id = self.sheet_id.get()
            count = self.count.get()
            is_date = self.flag_date.get()
            start_point = self.start_point.get()
            tool = utils.Excel2Img(file_path, dir_path, sheet_id)
            tool.excel2img(name_rule, start_point, count, is_date)
            messagebox.showinfo('提示', '转换成功！')
            # root.destroy()
        except Exception as e:
            print(e)
            messagebox.showinfo('提示', '转换失败！')

    def change_sheet(self, event=None):
        sheet_id = self.sheet_id.get()
        try:
            self.excel.set_sheet(sheet_id)
            self.preview()
        except Exception as err:
            print(err)

    def change_file(self):
        """
        设置可以选择的文件类型，不属于这个类型的，无法被选中
        """
        filetypes = [("Excel 工作簿", "*.xlsx"), ('Excel 启用宏的工作簿', '*.xlsm')]

        file_name = filedialog.askopenfilename(
            title='选择文件',
            filetypes=filetypes,
            initialdir='./'  # 打开当前程序工作目录
        )
        self.src_file.set(file_name)
        file_path = self.src_file.get()
        dir_path = self.dst_dir.get()
        sheet_id = self.sheet_id.get()
        self.excel = utils.Excel2Img(file_path, dir_path, sheet_id)
        self.preview()

    def change_dir(self):
        """
        设置导出目录
        """
        file_name = filedialog.askdirectory(
            title='选择文件夹',
            initialdir='./'  # 打开当前程序工作目录
        )
        self.dst_dir.set(file_name)


def main():
    try:
        MainWindow(
            version="0.1",
            update_time="2022.04.22",
            title="Etofile",
            icon_path="./icon/main.ico"
        )
    except Exception as err:
        print(err)


if __name__ == '__main__':
    main()
