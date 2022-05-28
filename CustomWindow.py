from ttkbootstrap import Window


class CustomWindow:
    """窗口模板"""

    def __init__(self, root: Window, **kwargs) -> None:
        """初始化"""
        super().__init__()

        self.root = root

        self.root.withdraw()

        """传参"""
        # self.theme_name = kwargs["theme_name"] if "theme_name" in kwargs else "default"  # 主题
        self.title = kwargs["title"] if "title" in kwargs else "无标题"  # 标题
        self.icon_path = kwargs["icon_path"] if "icon_path" in kwargs else ""  # 图标路径
        self.width = kwargs["width"] if "width" in kwargs else 0  # 宽
        self.height = kwargs["height"] if "height" in kwargs else 0  # 高

        #  窗体基本设置
        self.root.title(self.title)
        self.root.iconbitmap(self.icon_path)

        if 0 in (self.width, self.height):
            self.root.resizable(False, False)
        else:
            self.root.config(width=self.width, height=self.height)
