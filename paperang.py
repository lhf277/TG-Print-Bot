import os
import logging
import asyncio
import threading
import queue
import time
import textwrap
from telegram import Update
from telegram.ext import ApplicationBuilder, ContextTypes, CommandHandler, MessageHandler, filters
from telegram.error import NetworkError
import win32print
import win32ui
import win32con
from PIL import Image, ImageDraw, ImageFont, ImageOps

# 配置日志
# (已移动到下方配置区域后)

# 全局配置
BOT_TOKEN = ""
PRINTER_NAME = win32print.GetDefaultPrinter() # 默认使用系统默认打印机

# ================= 配置区域 =================
MAX_TEXT_LENGTH = 1500 # 字数限制 (可在此处修改)
LOG_LEVEL = logging.INFO # 日志等级: DEBUG, INFO, WARNING, ERROR
# ===========================================

# 配置日志
logging.basicConfig(
    format='%(asctime)s - %(levelname)s - %(message)s', # 简化格式，去掉 logger name
    level=LOG_LEVEL
)

# 屏蔽第三方库的繁琐日志
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("httpcore").setLevel(logging.WARNING)
logging.getLogger("telegram").setLevel(logging.WARNING)

PRINTER_WIDTH = 384 # 打印机宽度（像素），58mm通常为384，80mm通常为576

# 打印队列
print_queue = queue.Queue()

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text="你好！我是热敏打印机 Bot。\n发送文字或图片给我，我会帮你打印出来。"
    )

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    user = update.effective_user
    username = f"@{user.username}" if user.username else "No Username"
    nickname = user.full_name
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    
    if len(text) > MAX_TEXT_LENGTH:
        await context.bot.send_message(
            chat_id=update.effective_chat.id,
            text=f"❌ 你药剂吧干啥！文字太长了！请限制在 {MAX_TEXT_LENGTH} 字以内。"
        )
        return

    # 加入队列
    position = print_queue.qsize() + 1
    print_queue.put({
        'type': 'text', 
        'content': text, 
        'user': nickname,
        'header_info': {
            'username': username,
            'nickname': nickname,
            'time': timestamp
        }
    })
    
    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text=f"✅ 文字已加入打印队列。\n当前排队位置: {position}"
    )

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    username = f"@{user.username}" if user.username else "No Username"
    nickname = user.full_name
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())

    photo_file = await update.message.photo[-1].get_file()
    
    # 下载图片到临时文件
    file_path = f"temp_{int(time.time())}_{update.message.id}.jpg"
    await photo_file.download_to_drive(file_path)
    
    # 加入队列
    position = print_queue.qsize() + 1
    print_queue.put({
        'type': 'photo', 
        'content': file_path, 
        'user': nickname,
        'header_info': {
            'username': username,
            'nickname': nickname,
            'time': timestamp
        }
    })
    
    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text=f"✅ 图片已加入打印队列。\n当前排队位置: {position}"
    )

async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    """处理 Bot 运行时的错误"""
    # 打印简短的错误信息，避免刷屏
    if isinstance(context.error, NetworkError):
        logging.warning(f"网络连接不稳定: {context.error}")
        print("⚠️ 检测到网络连接断开，正在尝试自动重连...")
        
        # 停止当前应用运行
        # 注意：stop_running() 是一个普通方法，不是协程，不需要 await
        # 它会设置内部标志位，让 updater 在下一次循环时停止
        context.application.stop_running()
        
        # 为了确保立即退出，我们可以抛出一个系统退出异常
        # 但 stop_running 通常足够了，只要我们不阻塞
    else:
        logging.error("未处理的异常:", exc_info=context.error)

def text_to_image(text, header_info=None):
    """将文本转换为图片"""
    font_size = 24
    try:
        # 尝试使用微软雅黑，如果不存在则使用默认
        font = ImageFont.truetype("msyh.ttc", font_size)
        header_font = ImageFont.truetype("msyh.ttc", 20)
    except:
        font = ImageFont.load_default()
        header_font = ImageFont.load_default()

    # 准备头部信息
    header_lines = []
    if header_info:
        header_lines.append(f"User: {header_info['nickname']} ({header_info['username']})")
        header_lines.append(f"Time: {header_info['time']}")
        header_lines.append("-" * 30)

    # 自动换行
    lines = []
    # 粗略估算每行字符数 (PRINTER_WIDTH / (font_size / 2)) - 中文占全角
    # 这里使用 textwrap 简单处理，更精确的处理需要计算 textsize
    # 为了简单，我们假设一行大约 16 个中文字符 (384 / 24 = 16)
    chars_per_line = int(PRINTER_WIDTH / font_size) 
    
    for line in text.split('\n'):
        lines.extend(textwrap.wrap(line, width=chars_per_line))

    # 计算图片高度
    line_height = font_size + 4
    header_height = len(header_lines) * (20 + 4) + 10 if header_lines else 0
    img_height = len(lines) * line_height + 20 + header_height # 加上一些padding
    
    image = Image.new('RGB', (PRINTER_WIDTH, img_height), color='white')
    draw = ImageDraw.Draw(image)
    
    y = 10
    
    # 绘制头部
    for line in header_lines:
        draw.text((0, y), line, font=header_font, fill='black')
        y += 24

    # 绘制正文
    for line in lines:
        draw.text((0, y), line, font=font, fill='black')
        y += line_height
        
    temp_filename = f"temp_text_{int(time.time())}.jpg"
    image.save(temp_filename)
    return temp_filename

def print_image_file(file_path):
    """调用 Windows API 打印图片"""
    try:
        img = Image.open(file_path)
        
        # 调整图片宽度以适应打印机
        # 统一调整宽度到 PRINTER_WIDTH，确保填满纸张
        if img.width != PRINTER_WIDTH:
            ratio = PRINTER_WIDTH / img.width
            new_height = int(img.height * ratio)
            img = img.resize((PRINTER_WIDTH, new_height), Image.Resampling.LANCZOS)
        
        hDC = win32ui.CreateDC()
        hDC.CreatePrinterDC(PRINTER_NAME)
        
        hDC.StartDoc(f"Telegram Print Job {os.path.basename(file_path)}")
        
        # 获取打印机页面高度
        page_height = hDC.GetDeviceCaps(win32con.VERTRES)
        
        # 如果获取到的高度太小（异常），或者为0，给一个默认值（例如 3000 像素）
        if page_height <= 0:
            page_height = 3000
            
        # 打印逻辑：切片分页
        total_height = img.height
        current_y = 0
        
        from PIL import ImageWin
        
        while current_y < total_height:
            hDC.StartPage()
            
            remaining_height = total_height - current_y
            print_height = min(remaining_height, page_height)
            
            # 裁剪当前页的内容
            box = (0, current_y, img.width, current_y + print_height)
            region = img.crop(box)
            
            # 转换为 DIB
            dib = ImageOps.fit(region, (region.width, region.height))
            if dib.mode != 'RGB':
                dib = dib.convert('RGB')
            dib = ImageWin.Dib(dib)
            
            # 绘制到设备上下文
            dib.draw(hDC.GetHandleOutput(), (0, 0, region.width, region.height))
            
            hDC.EndPage()
            current_y += print_height
            
        hDC.EndDoc()
        hDC.DeleteDC()
        
    except Exception as e:
        logging.error(f"Windows 打印失败: {e}")
        raise e

def add_header_to_image(image_path, header_info):
    """给图片添加头部信息"""
    try:
        img = Image.open(image_path)
        
        # 准备头部信息
        header_lines = []
        if header_info:
            header_lines.append(f"User: {header_info['nickname']} ({header_info['username']})")
            header_lines.append(f"Time: {header_info['time']}")
            header_lines.append("-" * 30)
            
        if not header_lines:
            return image_path

        try:
            font = ImageFont.truetype("msyh.ttc", 20)
        except:
            font = ImageFont.load_default()
            
        header_height = len(header_lines) * 24 + 10
        
        # 创建新图片，宽度取原图和打印机宽度的最大值（通常我们会缩放到打印机宽度，所以这里取打印机宽度比较安全，但为了保持原图清晰度，我们先按原图宽处理，打印时再缩放）
        # 不过为了排版整齐，我们最好先调整原图宽度到 PRINTER_WIDTH，或者至少让头部宽度匹配
        
        target_width = PRINTER_WIDTH
        
        # 调整原图大小
        if img.width != target_width:
            ratio = target_width / img.width
            new_height = int(img.height * ratio)
            img = img.resize((target_width, new_height), Image.Resampling.LANCZOS)
            
        new_height = img.height + header_height
        new_img = Image.new('RGB', (target_width, new_height), color='white')
        draw = ImageDraw.Draw(new_img)
        
        # 绘制头部
        y = 5
        for line in header_lines:
            draw.text((5, y), line, font=font, fill='black')
            y += 24
            
        # 粘贴原图
        new_img.paste(img, (0, header_height))
        
        # 保存覆盖原文件
        new_img.save(image_path)
        return image_path
        
    except Exception as e:
        logging.error(f"添加图片头部失败: {e}")
        return image_path

def printer_worker():
    """后台线程：处理打印队列"""
    while True:
        try:
            task = print_queue.get()
            if task is None:
                break
            
            logging.info(f"正在处理来自 {task['user']} 的打印任务...")
            
            file_to_print = None
            header_info = task.get('header_info')
            
            if task['type'] == 'text':
                file_to_print = text_to_image(task['content'], header_info)
            elif task['type'] == 'photo':
                file_to_print = task['content']
                if header_info:
                    add_header_to_image(file_to_print, header_info)
            
            if file_to_print:
                print_image_file(file_to_print)
                
                # 清理文件
                try:
                    # 给一点时间释放句柄
                    time.sleep(1)
                    os.remove(file_to_print)
                except Exception as e:
                    logging.warning(f"清理文件失败: {e}")

            logging.info("打印完成")
            print_queue.task_done()
            
        except Exception as e:
            logging.error(f"打印出错: {e}")
            # 确保任务被标记为完成，避免队列阻塞
            try:
                print_queue.task_done()
            except:
                pass

if __name__ == '__main__':
    # 检查打印机
    try:
        default_printer = win32print.GetDefaultPrinter()
        print(f"当前默认打印机: {default_printer}")
        print("可用打印机列表:")
        for p in win32print.EnumPrinters(2):
            print(f" - {p[2]}")
        
        # 如果需要指定打印机，请修改全局变量 PRINTER_NAME
        if PRINTER_NAME != default_printer:
             print(f"注意：程序配置使用的打印机是 '{PRINTER_NAME}'，而非默认打印机。")

    except Exception as e:
        print(f"获取打印机信息失败: {e}")
        print("请确保已安装打印机驱动。")

    # 启动打印线程
    worker_thread = threading.Thread(target=printer_worker, daemon=True)
    worker_thread.start()
    
    # 启动 Bot
    if BOT_TOKEN == "YOUR_BOT_TOKEN_HERE":
        print("\n⚠️  请注意：你需要设置 BOT_TOKEN 才能运行！")
        print("请打开 c:\\paperang\\paperang.py 修改 BOT_TOKEN 变量。")
    else:
        while True:
            try:
                # 创建新的事件循环
                loop = asyncio.new_event_loop()
                asyncio.set_event_loop(loop)
                
                application = ApplicationBuilder().token(BOT_TOKEN).build()
                
                start_handler = CommandHandler('start', start)
                text_handler = MessageHandler(filters.TEXT & (~filters.COMMAND), handle_text)
                photo_handler = MessageHandler(filters.PHOTO, handle_photo)
                
                application.add_handler(start_handler)
                application.add_handler(text_handler)
                application.add_handler(photo_handler)
                application.add_error_handler(error_handler)
                
                print("Bot 已启动... 按 Ctrl+C 停止")
                
                # 使用 run_polling 的 close_loop=False 参数，避免它关闭循环
                # 但 Application.run_polling 内部处理比较复杂
                # 更稳健的方法是直接运行，如果抛出异常，确保清理
                
                application.run_polling()
                
            except Exception as e:
                # 只有在非网络错误时才打印详细堆栈，或者根据需要简化
                if "NetworkError" not in str(e) and "RemoteProtocolError" not in str(e):
                     logging.error(f"Bot 运行出错: {e}")
                else:
                     print(f"Bot 连接断开，5秒后自动重启...")
                
                time.sleep(5)
            except KeyboardInterrupt:
                print("用户停止程序")
                break
            finally:
                # 确保清理资源
                try:
                    loop = asyncio.get_event_loop()
                    if not loop.is_closed():
                        loop.close()
                except:
                    pass
