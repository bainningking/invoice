"""
发票识别工具
使用OCR技术从PDF发票中提取关键信息

主要功能：
1. 从PDF文件中提取右上角区域
2. 使用OCR技术识别文本
3. 从识别的文本中提取发票字段
4. 将结果导出到Excel文件
"""

import io
import logging
import os
import re
import subprocess
from datetime import datetime
from logging import Logger
from pathlib import Path
from typing import Dict, List, Optional

import fitz  # PyMuPDF
import flet as ft
import numpy as np
import pandas as pd
from paddleocr import PaddleOCR
from PIL import Image

os.environ['PADDLEOCR_HOME'] = '.'
# 常量定义
class Constants:
    """
    应用常量类

    包含应用中使用的所有常量，便于统一管理和修改
    """

    # PDF处理常量
    CROP_LEFT_RATIO = 0.7  # 左边界比例
    CROP_TOP_RATIO = 0.0  # 上边界比例
    CROP_RIGHT_RATIO = 1.0  # 右边界比例
    CROP_BOTTOM_RATIO = 0.2  # 下边界比例
    DEFAULT_DPI = 200  # 默认DPI

    # UI常量
    WINDOW_WIDTH = 800  # 窗口宽度
    WINDOW_HEIGHT = 700  # 窗口高度
    CONTAINER_WIDTH = 750  # 容器宽度
    CONTAINER_HEIGHT = 650  # 容器高度
    LOG_MAX_LINES = 10  # 日志最大行数
    LOG_HEIGHT = 200  # 日志区域高度

    # 发票字段常量
    INVOICE_FIELDS = ["页码", "发票代码", "发票号码", "开票日期", "校验码"]

    # 文件名常量
    CROPPED_IMAGES_DIR = "裁剪图片"  # 裁剪图片目录名
    EXCEL_FILE_PREFIX = "发票信息"  # Excel文件前缀

    # 错误消息
    ERROR_NO_FILE = "请先选择PDF文件"  # 未选择文件错误
    ERROR_FILE_NOT_EXIST = "PDF文件不存在"  # 文件不存在错误
    ERROR_PROCESSING_FAILED = "处理失败: 未识别到发票信息"  # 处理失败错误


# 配置日志记录器
def setup_logger(name: str = "invoice_processor") -> Logger:
    """
    设置并返回配置好的日志记录器

    Args:
        name: 日志记录器名称

    Returns:
        配置好的日志记录器实例
    """
    logger = logging.getLogger(name)
    if not logger.handlers:
        logger.setLevel(logging.INFO)

        # 创建控制台处理器
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)

        # 创建格式化器
        formatter = logging.Formatter(
            "[%(asctime)s] [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
        )
        console_handler.setFormatter(formatter)

        # 添加处理器到日志记录器
        logger.addHandler(console_handler)

    return logger


# 全局日志记录器实例
logger = setup_logger()


class FileUtils:
    """
    文件操作工具类

    提供常用的文件操作方法，包括目录创建、文件检查和系统打开文件等功能
    """

    @staticmethod
    def ensure_directory_exists(directory: str) -> bool:
        """
        确保目录存在

        Args:
            directory: 目录路径

        Returns:
            bool: 操作是否成功
        """
        try:
            os.makedirs(directory, exist_ok=True)
            return True
        except Exception as e:
            logger.error(f"创建目录失败: {directory}, 错误: {e}")
            return False

    @staticmethod
    def file_exists(file_path: str) -> bool:
        """
        检查文件是否存在

        Args:
            file_path: 文件路径

        Returns:
            bool: 文件是否存在
        """
        return os.path.exists(file_path)

    @staticmethod
    def open_file_with_system(file_path: str) -> bool:
        """
        使用系统默认程序打开文件

        Args:
            file_path: 文件路径

        Returns:
            bool: 操作是否成功
        """
        try:
            import platform

            if platform.system() == "Windows":
                os.startfile(file_path)
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", file_path])
            else:  # Linux
                subprocess.run(["xdg-open", file_path])
            return True
        except Exception as e:
            logger.error(f"打开文件失败: {file_path}, 错误: {e}")
            return False


class ErrorHandler:
    """
    错误处理工具类

    提供统一的错误处理方法，包括错误记录、安全执行和条件验证等功能
    """

    @staticmethod
    def handle_error(error: Exception, context: str, reraise: bool = False) -> None:
        """
        统一处理错误

        Args:
            error: 异常对象
            context: 错误上下文描述
            reraise: 是否重新抛出异常
        """
        error_msg = f"{context}: {str(error)}"
        logger.error(error_msg)
        if reraise:
            raise error

    @staticmethod
    def safe_execute(func, context: str, default_return=None, reraise: bool = False):
        """
        安全执行函数，捕获异常

        Args:
            func: 要执行的函数
            context: 执行上下文描述
            default_return: 异常时的默认返回值
            reraise: 是否重新抛出异常

        Returns:
            函数执行结果或默认返回值
        """
        try:
            return func()
        except Exception as e:
            ErrorHandler.handle_error(e, context, reraise)
            return default_return

    @staticmethod
    def log_and_continue(error: Exception, context: str, message: str = None):
        """
        记录错误但继续执行

        Args:
            error: 异常对象
            context: 错误上下文描述
            message: 额外的信息消息
        """
        error_msg = f"{context}: {str(error)}"
        logger.warning(error_msg)
        if message:
            logger.info(message)

    @staticmethod
    def validate_condition(condition: bool, error_message: str) -> bool:
        """
        验证条件，如果为False则记录错误

        Args:
            condition: 要验证的条件
            error_message: 验证失败时的错误消息

        Returns:
            bool: 条件是否为真
        """
        if not condition:
            logger.error(error_message)
            return False
        return True


class PDFProcessor:
    """
    PDF处理器类，负责PDF相关操作

    提供PDF页数获取、页面区域提取等功能
    """

    @staticmethod
    def get_page_count(pdf_path: str) -> int:
        """获取PDF页数"""
        try:
            doc = fitz.open(pdf_path)
            page_count = len(doc)
            doc.close()
            return page_count
        except Exception as e:
            logger.error(f"获取PDF页数失败: {e}")
            return 0

    @staticmethod
    def extract_page_region(
        pdf_path: str, page_num: int, save_image: bool = False, output_dir: str = None
    ) -> Optional[Image.Image]:
        """
        从PDF页面中提取右上角区域
        Args:
            pdf_path: PDF文件路径
            page_num: 页码（从0开始）
            save_image: 是否保存裁剪后的图片
            output_dir: 输出目录
        Returns:
            提取的图像，如果失败返回None
        """
        try:
            logger.info(f"开始处理PDF第{page_num + 1}页")

            # 打开PDF文件
            doc = fitz.open(pdf_path)
            if page_num >= len(doc):
                logger.error(f"页码{page_num}超出范围，PDF共{len(doc)}页")
                return None

            page = doc[page_num]

            # 获取页面尺寸
            rect = page.rect
            page_width = rect.width
            page_height = rect.height

            logger.info(f"页面尺寸: {page_width} x {page_height}")

            # 计算右上角区域坐标（横向70-100%，纵向0-25%）
            left = page_width * Constants.CROP_LEFT_RATIO
            top = page_height * Constants.CROP_TOP_RATIO
            right = page_width * Constants.CROP_RIGHT_RATIO
            bottom = page_height * Constants.CROP_BOTTOM_RATIO

            crop_rect = fitz.Rect(left, top, right, bottom)
            logger.debug(
                f"裁剪区域: left={left:.2f}, top={top:.2f}, right={right:.2f}, bottom={bottom:.2f}"
            )

            # 裁剪页面并转换为图像
            pix = page.get_pixmap(clip=crop_rect, dpi=Constants.DEFAULT_DPI)
            img_data = pix.tobytes("png")
            img = Image.open(io.BytesIO(img_data))

            # 保存裁剪后的图片
            if save_image and output_dir:

                def save_image():
                    # 确保输出目录存在
                    FileUtils.ensure_directory_exists(output_dir)

                    # 生成文件名
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"cropped_page_{page_num + 1}_{timestamp}.png"
                    filepath = os.path.join(output_dir, filename)

                    # 保存图片
                    img.save(filepath)
                    logger.info(f"裁剪图片已保存到: {filepath}")

                ErrorHandler.safe_execute(save_image, "保存裁剪图片")

            doc.close()
            logger.info(f"第{page_num + 1}页区域提取完成")
            return img

        except Exception as e:
            logger.error(f"提取第{page_num + 1}页区域失败: {e}")
            return None


class OCREngine:
    """OCR引擎类，负责文字识别"""

    def __init__(self):
        logger.info("初始化OCR引擎")
        # 初始化PaddleOCR
        try:
            self.ocr = PaddleOCR(use_angle_cls=True, lang="ch")
            logger.info("PaddleOCR初始化成功")
        except Exception as e:
            logger.error(f"PaddleOCR初始化失败: {e}")
            raise

    def recognize_text(self, image: Image.Image) -> str:
        """
        使用PaddleOCR识别图像中的文本
        Args:
            image: PIL图像对象
        Returns:
            识别的文本
        """
        try:
            logger.info("开始OCR识别")

            # 将PIL图像转换为numpy数组
            img_array = np.array(image)

            # 使用PaddleOCR识别 - 使用推荐的API
            try:
                # 使用predict API（推荐）
                result = self.ocr.predict(img_array)
                logger.debug("使用predict API")
            except Exception as e1:
                logger.warning(f"predict API失败，尝试ocr API: {e1}")
                # 尝试旧的ocr API作为备选
                result = self.ocr.ocr(img_array)
                logger.debug("使用ocr API")

            # 提取文本 - 处理不同的返回格式
            texts = []
            if (
                isinstance(result, list)
                and len(result) > 0
                and isinstance(result[0], dict)
            ):
                # 新版PaddleOCR返回格式（字典列表）
                for page_result in result:
                    if "rec_texts" in page_result:
                        texts.extend(page_result["rec_texts"])
                logger.debug(f"新版PaddleOCR返回格式，识别到{len(texts)}个文本")
            elif isinstance(result, dict) and "texts" in result:
                # 另一种新API返回格式
                texts = result["texts"]
                logger.debug(f"新API返回格式，识别到{len(texts)}个文本")
            elif isinstance(result, list):
                # 旧API返回格式
                for line in result:
                    if line:
                        for word_info in line:
                            if (
                                isinstance(word_info, (list, tuple))
                                and len(word_info) >= 2
                            ):
                                if word_info[1] and word_info[1][0]:  # 确保文本不为空
                                    texts.append(word_info[1][0])
                            elif isinstance(word_info, str):
                                texts.append(word_info)
                logger.debug(f"旧API返回格式，识别到{len(texts)}个文本")
            else:
                logger.warning(f"未知的返回格式: {type(result)}")

            recognized_text = "\n".join(texts)
            logger.info(f"OCR识别完成，识别到文本: {recognized_text}")
            return recognized_text

        except Exception as e:
            logger.error(f"OCR识别失败: {e}")
            return ""


class InvoiceFieldExtractor:
    """发票字段提取器类，负责从文本中提取发票信息"""

    @staticmethod
    def extract_fields(text: str) -> Dict[str, str]:
        """
        从识别的文本中提取发票字段
        Args:
            text: OCR识别的文本
        Returns:
            包含发票字段的字典
        """
        logger.info("开始提取发票字段")

        fields = {field: "" for field in Constants.INVOICE_FIELDS if field != "页码"}

        try:
            # 使用正则表达式提取各个字段
            lines = text.split("\n")

            for line in lines:
                line = line.strip()
                if not line:
                    continue

                # 发票代码
                if "发票代码" in line:
                    code_match = re.search(r"发票代码[：:\s]*(\d+)", line)
                    if code_match:
                        fields["发票代码"] = code_match.group(1)
                        logger.debug(f"提取发票代码: {fields['发票代码']}")

                # 发票号码
                elif "发票号码" in line:
                    number_match = re.search(r"发票号码[：:\s]*(\d+)", line)
                    if number_match:
                        fields["发票号码"] = number_match.group(1)
                        logger.debug(f"提取发票号码: {fields['发票号码']}")

                # 开票日期
                elif "开票日期" in line:
                    date_match = re.search(
                        r"开票日期[：:\s]*(\d{4}年\d{1,2}月\d{1,2}日)", line
                    )
                    if date_match:
                        fields["开票日期"] = date_match.group(1)
                        logger.debug(f"提取开票日期: {fields['开票日期']}")

                # 校验码
                elif "校验码" in line:
                    code_match = re.search(r"校验码[：:\s]*([0-9A-Za-z\s]+)", line)
                    if code_match:
                        fields["校验码"] = code_match.group(1).strip()
                        logger.debug(f"提取校验码: {fields['校验码']}")

            logger.debug(f"字段提取完成: {fields}")
            return fields

        except Exception as e:
            logger.error(f"字段提取失败: {e}")
            return fields


class InvoiceProcessor:
    """发票处理器类，协调各个组件完成发票处理"""

    def __init__(self):
        logger.info("初始化发票处理器")
        self.pdf_processor = PDFProcessor()
        self.ocr_engine = OCREngine()
        self.field_extractor = InvoiceFieldExtractor()

    def process_pdf(
        self,
        pdf_path: str,
        progress_callback=None,
        save_cropped_images: bool = False,
        output_dir: str = None,
    ) -> List[Dict[str, str]]:
        """
        处理整个PDF文件
        Args:
            pdf_path: PDF文件路径
            progress_callback: 进度回调函数
            save_cropped_images: 是否保存裁剪后的图片
            output_dir: 输出目录
        Returns:
            每页的发票信息列表
        """
        logger.info(f"开始处理PDF文件: {pdf_path}")

        try:
            # 获取PDF页数
            total_pages = self.pdf_processor.get_page_count(pdf_path)
            if total_pages == 0:
                logger.error("无法获取PDF页数或PDF为空")
                return []

            logger.info(f"PDF文件共{total_pages}页")

            results = []

            # 创建图片输出目录
            images_dir = None
            if save_cropped_images and output_dir:
                images_dir = os.path.join(output_dir, Constants.CROPPED_IMAGES_DIR)
                os.makedirs(images_dir, exist_ok=True)
                logger.info(f"裁剪图片将保存到: {images_dir}")

            for page_num in range(total_pages):
                if progress_callback:
                    progress_callback(f"正在处理第{page_num + 1}/{total_pages}页...")

                # 提取页面区域
                image = self.pdf_processor.extract_page_region(
                    pdf_path, page_num, save_cropped_images, images_dir
                )
                if image is None:
                    logger.error(f"第{page_num + 1}页区域提取失败")
                    continue

                # OCR识别
                text = self.ocr_engine.recognize_text(image)
                if not text:
                    logger.error(f"第{page_num + 1}页OCR识别失败")
                    continue

                # 提取字段
                fields = self.field_extractor.extract_fields(text)
                fields["页码"] = page_num + 1

                results.append(fields)
                logger.info(f"第{page_num + 1}页处理完成")

            logger.info(f"PDF处理完成，共处理{len(results)}页")
            return results

        except Exception as e:
            logger.error(f"PDF处理失败: {e}")
            return []


class UIComponents:
    """UI组件类，负责创建和管理界面元素"""

    @staticmethod
    def create_button(
        text: str,
        icon,
        on_click,
        color: str = ft.Colors.BLUE_600,
        disabled: bool = False,
    ) -> ft.ElevatedButton:
        """创建标准按钮"""
        return ft.ElevatedButton(
            text,
            icon=icon,
            on_click=on_click,
            disabled=disabled,
            style=ft.ButtonStyle(
                shape=ft.RoundedRectangleBorder(radius=12),
                padding=ft.padding.symmetric(horizontal=20, vertical=15),
                bgcolor=color,
                elevation=3,
            ),
        )

    @staticmethod
    def create_container(
        content, padding: int = 20, margin=None, width=None, height=None
    ) -> ft.Container:
        """创建标准容器"""
        return ft.Container(
            content=content,
            padding=padding,
            bgcolor=ft.Colors.WHITE,
            border_radius=15,
            shadow=ft.BoxShadow(
                blur_radius=10, spread_radius=1, color=ft.Colors.GREY_300
            ),
            margin=margin,
            width=width,
            height=height,
        )

    @staticmethod
    def create_section(title: str, content) -> ft.Container:
        """创建标准区域"""
        return UIComponents.create_container(
            content=ft.Column(
                [
                    ft.Text(
                        title,
                        size=16,
                        weight=ft.FontWeight.BOLD,
                        color=ft.Colors.BLUE_600,
                    ),
                    content,
                ],
                spacing=10,
            ),
            margin=ft.margin.only(bottom=15),
        )


class InvoiceBusinessLogic:
    """发票处理业务逻辑类"""

    def __init__(self):
        self.processor = None
        self.last_excel_path = ""

    def validate_file(self, pdf_path: str) -> bool:
        """验证PDF文件"""
        if not pdf_path:
            logger.error("未选择PDF文件")
            return False

        if not FileUtils.file_exists(pdf_path):
            logger.error(f"PDF文件不存在: {pdf_path}")
            return False

        return True

    def initialize_processor(self) -> bool:
        """初始化处理器"""

        def init():
            if self.processor is None:
                self.processor = InvoiceProcessor()
            return True

        return ErrorHandler.safe_execute(init, "初始化处理器", False)

    def process_invoice_file(
        self, pdf_path: str, save_images: bool, output_path: str, progress_callback
    ) -> List[Dict[str, str]]:
        """处理发票文件"""
        if not self.validate_file(pdf_path):
            return []

        if not self.initialize_processor():
            return []

        # 处理PDF
        results = self.processor.process_pdf(
            pdf_path,
            progress_callback,
            save_cropped_images=save_images,
            output_dir=output_path,
        )

        return results

    def export_results(self, results: List[Dict[str, str]], output_path: str) -> str:
        """导出结果到Excel"""
        logger.info("开始导出Excel")

        # 生成文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{Constants.EXCEL_FILE_PREFIX}_{timestamp}.xlsx"
        filepath = os.path.join(output_path, filename)

        # 创建DataFrame
        df = pd.DataFrame(results)

        # 重新排列列的顺序
        existing_columns = [
            col for col in Constants.INVOICE_FIELDS if col in df.columns
        ]
        df = df[existing_columns]

        # 导出到Excel
        df.to_excel(filepath, index=False, engine="openpyxl")

        logger.info(f"Excel导出完成: {filepath}")
        self.last_excel_path = filepath
        return filepath

    def open_output_file(self) -> bool:
        """打开输出文件"""
        if not self.last_excel_path:
            logger.error("输出文件不存在或尚未生成")
            return False

        if not FileUtils.file_exists(self.last_excel_path):
            logger.error("输出文件不存在或尚未生成")
            return False

        if FileUtils.open_file_with_system(self.last_excel_path):
            logger.info(f"已打开文件: {self.last_excel_path}")
            return True

        return False


class InvoiceApp:
    """发票识别GUI应用"""

    def __init__(self):
        logger.info("初始化GUI应用")
        self.pdf_path = ""
        self.output_path = str(Path(__file__).parent)
        self.business_logic = InvoiceBusinessLogic()
        self.file_picker = None
        self.dir_picker = None
        self.save_images_checkbox = None
        self.process_btn = None
        self.is_processing = False
        self.log_container = None
        self.log_list = []
        self.max_log_lines = Constants.LOG_MAX_LINES
        self.page = None  # 页面对象引用

    def main(self, page: ft.Page):
        """主界面"""
        self.page = page  # 保存页面引用
        page.title = "发票识别工具"
        page.window.width = Constants.WINDOW_WIDTH
        page.window.height = Constants.WINDOW_HEIGHT
        page.theme_mode = ft.ThemeMode.LIGHT
        page.vertical_alignment = ft.MainAxisAlignment.CENTER
        page.horizontal_alignment = ft.CrossAxisAlignment.CENTER

        # 设置页面背景色
        page.bgcolor = ft.Colors.GREY_50

        # 初始化文件选择器
        self.file_picker = ft.FilePicker(on_result=self.on_file_selected)
        self.dir_picker = ft.FilePicker(on_result=self.on_directory_selected)

        # 将文件选择器添加到页面覆盖层
        page.overlay.extend([self.file_picker, self.dir_picker])

        # 文件路径显示（可复制）
        self.pdf_file_text = ft.Text(
            "请选择PDF文件", size=14, color=ft.Colors.GREY_600, selectable=True
        )
        self.output_path_text = ft.Text(
            f"输出位置: {self.output_path}",
            size=14,
            color=ft.Colors.GREY_600,
            selectable=True,
        )

        # 日志显示区域
        self.log_container = ft.Column(
            controls=[
                ft.Text(
                    "等待开始处理...",
                    size=12,
                    color=ft.Colors.BLUE_600,
                    selectable=True,
                )
            ],
            scroll=ft.ScrollMode.AUTO,
            height=Constants.LOG_HEIGHT,
            spacing=2,
        )

        # 添加保存裁剪图片的选项
        self.save_images_checkbox = ft.Checkbox(
            label="保存裁剪后的图片（用于调试）",
            value=False,
            on_change=self.on_save_images_change,
        )

        # 按钮
        upload_btn = UIComponents.create_button(
            "上传PDF", ft.Icons.UPLOAD_FILE, self.upload_pdf
        )
        output_btn = UIComponents.create_button(
            "输出位置", ft.Icons.FOLDER_OPEN, self.select_output
        )
        self.process_btn = UIComponents.create_button(
            "开始处理",
            ft.Icons.PLAY_ARROW,
            self.process_invoices,
            ft.Colors.GREEN_600,
            disabled=True,
        )

        # 复制按钮
        copy_path_btn = ft.IconButton(
            icon=ft.Icons.COPY,
            tooltip="复制路径",
            on_click=self.copy_path,
            icon_size=16,
        )

        copy_log_btn = ft.IconButton(
            icon=ft.Icons.COPY_ALL,
            tooltip="复制日志",
            on_click=self.copy_log,
            icon_size=16,
        )

        # 打开输出文件按钮
        self.open_file_btn_ref = UIComponents.create_button(
            "打开输出文件",
            ft.Icons.OPEN_IN_NEW,
            self.open_output_file,
            ft.Colors.ORANGE_600,
            disabled=True,
        )
        open_file_btn = self.open_file_btn_ref

        # 布局
        page.add(
            ft.Container(
                content=ft.Column(
                    [
                        # 标题
                        ft.Container(
                            content=ft.Text(
                                "发票识别工具",
                                size=28,
                                weight=ft.FontWeight.BOLD,
                                color=ft.Colors.BLUE_700,
                            ),
                            margin=ft.margin.only(bottom=20),
                        ),
                        # 文件选择区域
                        ft.Container(
                            content=ft.Column(
                                [
                                    ft.Text(
                                        "选择文件",
                                        size=16,
                                        weight=ft.FontWeight.BOLD,
                                        color=ft.Colors.BLUE_600,
                                    ),
                                    ft.Row(
                                        [
                                            upload_btn,
                                            ft.Container(
                                                content=self.pdf_file_text,
                                                expand=True,
                                                margin=ft.margin.only(left=10),
                                            ),
                                            copy_path_btn,
                                        ],
                                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                    ),
                                ],
                                spacing=10,
                            ),
                            padding=20,
                            bgcolor=ft.Colors.WHITE,
                            border_radius=15,
                            shadow=ft.BoxShadow(
                                blur_radius=10,
                                spread_radius=1,
                                color=ft.Colors.GREY_300,
                            ),
                            margin=ft.margin.only(bottom=15),
                        ),
                        # 输出位置区域
                        ft.Container(
                            content=ft.Column(
                                [
                                    ft.Text(
                                        "输出位置",
                                        size=16,
                                        weight=ft.FontWeight.BOLD,
                                        color=ft.Colors.BLUE_600,
                                    ),
                                    ft.Row(
                                        [
                                            output_btn,
                                            ft.Container(
                                                content=self.output_path_text,
                                                expand=True,
                                                margin=ft.margin.only(left=10),
                                            ),
                                            copy_path_btn,
                                        ],
                                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                    ),
                                ],
                                spacing=10,
                            ),
                            padding=20,
                            bgcolor=ft.Colors.WHITE,
                            border_radius=15,
                            shadow=ft.BoxShadow(
                                blur_radius=10,
                                spread_radius=1,
                                color=ft.Colors.GREY_300,
                            ),
                            margin=ft.margin.only(bottom=15),
                        ),
                        # 选项区域
                        ft.Container(
                            content=self.save_images_checkbox,
                            padding=20,
                            bgcolor=ft.Colors.WHITE,
                            border_radius=15,
                            shadow=ft.BoxShadow(
                                blur_radius=10,
                                spread_radius=1,
                                color=ft.Colors.GREY_300,
                            ),
                            margin=ft.margin.only(bottom=15),
                        ),
                        # 处理按钮区域
                        ft.Container(
                            content=ft.Row(
                                [self.process_btn, open_file_btn], spacing=10
                            ),
                            margin=ft.margin.only(bottom=15),
                        ),
                        # 日志区域
                        ft.Container(
                            content=ft.Column(
                                [
                                    ft.Row(
                                        [
                                            ft.Text(
                                                "处理日志",
                                                size=16,
                                                weight=ft.FontWeight.BOLD,
                                                color=ft.Colors.BLUE_600,
                                            ),
                                            copy_log_btn,
                                        ],
                                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                    ),
                                    ft.Container(
                                        content=self.log_container,
                                        padding=10,
                                        bgcolor=ft.Colors.GREY_100,
                                        border_radius=10,
                                        border=ft.border.all(1, ft.Colors.GREY_300),
                                    ),
                                ],
                                spacing=10,
                            ),
                            padding=20,
                            bgcolor=ft.Colors.WHITE,
                            border_radius=15,
                            shadow=ft.BoxShadow(
                                blur_radius=10,
                                spread_radius=1,
                                color=ft.Colors.GREY_300,
                            ),
                            width=Constants.CONTAINER_WIDTH,
                        ),
                    ],
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                    scroll=ft.ScrollMode.AUTO,
                ),
                padding=30,
                bgcolor=ft.Colors.GREY_50,
                border_radius=20,
                width=Constants.CONTAINER_WIDTH,
                height=Constants.CONTAINER_HEIGHT,
            )
        )

        logger.info("GUI界面初始化完成")

    def upload_pdf(self, e):
        """上传PDF文件"""
        logger.info("用户点击上传PDF按钮")

        try:
            # 显示文件选择器
            self.file_picker.pick_files(
                dialog_title="选择PDF文件",
                allowed_extensions=["pdf"],
                initial_directory=str(Path.home()),
            )

        except Exception as ex:
            logger.error(f"文件选择失败: {ex}")
            self.update_progress(f"文件选择失败: {ex}")

    def on_file_selected(self, e: ft.FilePickerResultEvent):
        """文件选择完成回调"""
        logger.debug(f"文件选择回调触发: {e}")
        if e.files:
            self.pdf_path = e.files[0].path
            logger.info(f"用户选择文件: {self.pdf_path}")
            self.pdf_file_text.value = f"已选择: {Path(self.pdf_path).name}"
            self.pdf_file_text.color = ft.Colors.BLUE_600
            self.pdf_file_text.update()
            self.update_progress("PDF文件已选择，可以开始处理")
            self.update_process_button_state()
        else:
            logger.info("用户取消了文件选择")
            self.update_progress("未选择文件")
            self.update_process_button_state()

    def select_output(self, e):
        """选择输出位置"""
        logger.info("用户点击选择输出位置按钮")

        try:
            # 显示目录选择器
            self.dir_picker.get_directory_path(
                dialog_title="选择输出目录", initial_directory=self.output_path
            )

        except Exception as ex:
            logger.error(f"目录选择失败: {ex}")
            self.update_progress(f"目录选择失败: {ex}")

    def on_directory_selected(self, e: ft.FilePickerResultEvent):
        """目录选择完成回调"""
        logger.debug(f"目录选择回调触发: {e}")
        if e.path:
            self.output_path = e.path
            logger.info(f"用户选择输出目录: {self.output_path}")
            self.output_path_text.value = f"输出位置: {self.output_path}"
            self.output_path_text.color = ft.Colors.BLUE_600
            self.output_path_text.update()
            self.update_progress("输出位置已更新")
        else:
            logger.info("用户取消了目录选择")
            self.update_progress("未更改输出位置")

    def update_progress(self, message: str):
        """更新进度显示"""
        logger.debug(f"更新进度: {message}")

        # 添加时间戳
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}"

        # 添加到日志列表
        self.log_list.append(log_entry)

        # 限制日志行数
        if len(self.log_list) > self.max_log_lines:
            self.log_list = self.log_list[-self.max_log_lines :]

        # 更新日志显示
        self.log_container.controls = [
            ft.Text(
                log,
                size=12,
                selectable=True,
                color=ft.Colors.BLUE_700 if "错误" not in log else ft.Colors.RED_600,
            )
            for log in self.log_list
        ]
        self.log_container.update()

    def on_save_images_change(self, e):
        """保存图片选项变更回调"""
        logger.debug(f"保存裁剪图片选项: {self.save_images_checkbox.value}")

    def update_process_button_state(self):
        """更新处理按钮状态"""
        if self.process_btn:
            # 只有选择了PDF文件且不在处理中时才启用按钮
            self.process_btn.disabled = not self.pdf_path or self.is_processing
            self.process_btn.update()

    def copy_path(self, e):
        """复制路径到剪贴板"""

        def copy():
            if self.pdf_path:
                page_data = e.page
                page_data.set_clipboard(self.pdf_path)
                self.update_progress("文件路径已复制到剪贴板")
            else:
                page_data = e.page
                page_data.set_clipboard(self.output_path)
                self.update_progress("输出路径已复制到剪贴板")

        ErrorHandler.safe_execute(copy, "复制路径")

    def copy_log(self, e):
        """复制日志到剪贴板"""

        def copy():
            if self.log_list:
                log_text = "\n".join(self.log_list)
                page_data = e.page
                page_data.set_clipboard(log_text)
                self.update_progress("日志已复制到剪贴板")
            else:
                self.update_progress("没有日志可复制")

        ErrorHandler.safe_execute(copy, "复制日志")

    def process_invoices(self, e):
        """处理发票"""
        logger.info("用户点击开始处理按钮")

        # 设置处理状态
        self.is_processing = True
        self.update_process_button_state()

        try:
            self.update_progress("正在初始化处理器...")

            # 获取是否保存裁剪图片的选项
            save_images = (
                self.save_images_checkbox.value if self.save_images_checkbox else False
            )

            # 处理PDF
            results = self.business_logic.process_invoice_file(
                self.pdf_path, save_images, self.output_path, self.update_progress
            )

            if not results:
                logger.error("处理结果为空")
                self.update_progress(f"错误: {Constants.ERROR_PROCESSING_FAILED}")
                return

            # 导出到Excel
            self.update_progress("正在导出到Excel...")
            excel_path = self.business_logic.export_results(results, self.output_path)

            # 启用打开文件按钮
            self.enable_open_file_button()

            # 如果保存了裁剪图片，提示用户
            if save_images:
                images_dir = os.path.join(
                    self.output_path, Constants.CROPPED_IMAGES_DIR
                )
                self.update_progress(
                    f"处理完成！结果已保存到: {excel_path}，裁剪图片已保存到: {images_dir}"
                )
            else:
                self.update_progress(f"处理完成！结果已保存到: {excel_path}")

            logger.info(f"处理完成，结果已保存到: {excel_path}")

        except Exception as ex:
            logger.error(f"处理失败: {ex}")
            self.update_progress(f"处理失败: {ex}")
        finally:
            # 恢复处理状态
            self.is_processing = False
            self.update_process_button_state()

    def enable_open_file_button(self):
        """启用打开文件按钮"""
        # 直接通过保存的引用更新按钮状态
        if hasattr(self, "open_file_btn_ref"):
            self.open_file_btn_ref.disabled = False
            self.open_file_btn_ref.update()

    def open_output_file(self, e):
        """打开输出文件"""
        if self.business_logic.open_output_file():
            self.update_progress(f"已打开文件: {self.business_logic.last_excel_path}")
        else:
            self.update_progress("输出文件不存在或尚未生成")


def main():
    """
    主函数

    应用程序入口点，创建并启动发票识别应用
    """
    logger.info("启动发票识别应用")

    # 启动应用
    app = InvoiceApp()
    ft.app(target=app.main)


if __name__ == "__main__":
    main()
