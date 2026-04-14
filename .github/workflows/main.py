#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
海康威视自动化桌面版
功能：自动登录 → 入侵报警 → 详情 → 在线获取
"""

import asyncio
import json
import os
import sys
from pathlib import Path
from playwright.async_api import async_playwright
import tkinter as tk
from tkinter import messagebox, ttk
import threading

# 配置文件路径（打包后放在exe同目录）
def get_config_path():
    if getattr(sys, 'frozen', False):
        # 打包后的exe运行
        return Path(sys.executable).parent / "config.json"
    else:
        # 开发环境
        return Path(__file__).parent / "config.json"

CONFIG_FILE = get_config_path()

class HikvisionApp:
    def __init__(self):
        self.config = self.load_config()
        self.root = tk.Tk()
        self.root.title("海康威视自动化工具 v1.0")
        self.root.geometry("500x400")
        self.root.resizable(False, False)
        
        # 设置图标（如果有的话）
        try:
            self.root.iconbitmap(default='')
        except:
            pass
        
        # 居中窗口
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - 500) // 2
        y = (screen_height - 400) // 2
        self.root.geometry(f"500x400+{x}+{y}")
        
        self.setup_ui()
        
    def setup_ui(self):
        """设置界面"""
        # 标题
        title_frame = tk.Frame(self.root, bg="#1890ff", height=60)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)
        
        title = tk.Label(title_frame, text="海康威视综合安防管理平台", 
                        font=("Microsoft YaHei", 16, "bold"), 
                        fg="white", bg="#1890ff")
        title.pack(expand=True)
        
        # 主内容区
        content = tk.Frame(self.root, bg="white")
        content.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)
        
        # 状态显示
        tk.Label(content, text="当前状态:", font=("Microsoft YaHei", 11, "bold"), 
                bg="white", fg="#333").pack(anchor=tk.W)
        
        self.status_label = tk.Label(content, text="就绪 - 请点击下方按钮开始", 
                                    font=("Microsoft YaHei", 10), 
                                    fg="#666", bg="white",
                                    wraplength=440, justify=tk.LEFT)
        self.status_label.pack(anchor=tk.W, pady=(5, 15))
        
        # 进度条
        self.progress = ttk.Progressbar(content, length=440, mode='determinate', maximum=100)
        self.progress.pack(pady=10)
        
        # 按钮区域
        btn_frame = tk.Frame(content, bg="white")
        btn_frame.pack(pady=20)
        
        self.start_btn = tk.Button(btn_frame, text="▶ 开始执行", command=self.start_task,
                                  bg="#1890ff", fg="white", 
                                  font=("Microsoft YaHei", 12, "bold"),
                                  width=14, height=1, relief=tk.FLAT, 
                                  cursor="hand2", activebackground="#40a9ff")
        self.start_btn.pack(side=tk.LEFT, padx=10)
        
        tk.Button(btn_frame, text="⚙ 账号设置", command=self.show_settings,
                 font=("Microsoft YaHei", 12),
                 width=14, height=1, relief=tk.GROOVE,
                 bg="#f5f5f5", activebackground="#e8e8e8").pack(side=tk.LEFT, padx=10)
        
        # 日志区域
        log_frame = tk.Frame(content, bg="white")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(log_frame, text="执行日志:", font=("Microsoft YaHei", 9), 
                bg="white", fg="#999").pack(anchor=tk.W)
        
        self.log_text = tk.Text(log_frame, height=6, width=60, state=tk.DISABLED,
                               font=("Consolas", 9), bg="#f8f8f8", fg="#333",
                               relief=tk.FLAT, padx=5, pady=5)
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
        
        # 底部信息
        footer = tk.Label(self.root, text="首次使用请先设置账号密码", 
                         font=("Microsoft YaHei", 9), fg="#999")
        footer.pack(pady=10)
        
    def log(self, message):
        """添加日志"""
        import datetime
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        full_msg = f"[{timestamp}] {message}"
        
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{full_msg}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        
        self.status_label.config(text=message)
        self.root.update()
        
    def load_config(self):
        """加载配置"""
        default_config = {
            "url": "https://10.108.90.1",
            "username": "",
            "password": "",
            "headless": True,  # True=隐藏浏览器，False=显示
            "slow_mo": 500     # 操作延迟毫秒
        }
        
        if CONFIG_FILE.exists():
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    saved = json.load(f)
                    default_config.update(saved)
            except:
                pass
        
        return default_config
    
    def save_config(self, **kwargs):
        """保存配置"""
        self.config.update(kwargs)
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            messagebox.showerror("保存失败", f"无法保存配置：{e}")
            return False
    
    def show_settings(self):
        """显示设置窗口"""
        settings = tk.Toplevel(self.root)
        settings.title("账号设置")
        settings.geometry("450x350")
        settings.resizable(False, False)
        settings.transient(self.root)
        settings.grab_set()
        
        # 居中
        settings.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 450) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 350) // 2
        settings.geometry(f"+{x}+{y}")
        
        frame = tk.Frame(settings, padx=30, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # 平台地址
        tk.Label(frame, text="平台地址:", font=("Microsoft YaHei", 10)).pack(anchor=tk.W, pady=(0, 5))
        url_entry = tk.Entry(frame, width=45, font=("Microsoft YaHei", 10))
        url_entry.insert(0, self.config["url"])
        url_entry.pack(fill=tk.X, pady=(0, 15))
        
        # 用户名
        tk.Label(frame, text="用户名:", font=("Microsoft YaHei", 10)).pack(anchor=tk.W, pady=(0, 5))
        user_entry = tk.Entry(frame, width=45, font=("Microsoft YaHei", 10))
        user_entry.insert(0, self.config["username"])
        user_entry.pack(fill=tk.X, pady=(0, 15))
        
        # 密码
        tk.Label(frame, text="密码:", font=("Microsoft YaHei", 10)).pack(anchor=tk.W, pady=(0, 5))
        pwd_entry = tk.Entry(frame, width=45, show="●", font=("Microsoft YaHei", 10))
        pwd_entry.insert(0, self.config["password"])
        pwd_entry.pack(fill=tk.X, pady=(0, 15))
        
        # 显示浏览器选项
        show_var = tk.BooleanVar(value=not self.config["headless"])
        tk.Checkbutton(frame, text="执行时显示浏览器窗口（调试用）", 
                      variable=show_var, font=("Microsoft YaHei", 9)).pack(anchor=tk.W, pady=(0, 20))
        
        def save():
            success = self.save_config(
                url=url_entry.get().strip(),
                username=user_entry.get().strip(),
                password=pwd_entry.get(),
                headless=not show_var.get()
            )
            if success:
                messagebox.showinfo("成功", "设置已保存！", parent=settings)
                settings.destroy()
        
        tk.Button(frame, text="💾 保存设置", command=save,
                 bg="#1890ff", fg="white", font=("Microsoft YaHei", 11, "bold"),
                 width=15, relief=tk.FLAT, cursor="hand2").pack(pady=10)
    
    def start_task(self):
        """启动任务"""
        if not self.config["username"] or not self.config["password"]:
            messagebox.showwarning("提示", "请先设置账号密码！\n\n点击「账号设置」按钮进行配置。")
            self.show_settings()
            return
        
        self.start_btn.config(state=tk.DISABLED, text="执行中...", bg="#ccc")
        self.progress['value'] = 0
        
        # 在新线程运行（避免阻塞UI）
        thread = threading.Thread(target=self.run_async_task, daemon=True)
        thread.start()
    
    def run_async_task(self):
        """运行异步任务"""
        try:
            asyncio.run(self.execute_automation())
        except Exception as e:
            self.log(f"❌ 致命错误: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("错误", str(e)))
        finally:
            self.root.after(0, self.reset_ui)
    
    async def execute_automation(self):
        """执行自动化流程"""
        browser = None
        try:
            async with async_playwright() as p:
                self.update_progress(5, "正在启动浏览器...")
                
                browser = await p.chromium.launch(
                    headless=self.config["headless"],
                    args=[
                        '--ignore-certificate-errors',
                        '--ignore-certificate-errors-spki-list',
                        '--disable-web-security'
                    ],
                    slow_mo=self.config["slow_mo"]
                )
                
                context = await browser.new_context(
                    viewport={'width': 1920, 'height': 1080},
                    ignore_https_errors=True
                )
                
                page = await context.new_page()
                
                # 1. 登录
                self.update_progress(15, "正在访问登录页面...")
                await page.goto(self.config["url"], timeout=30000)
                await asyncio.sleep(2)
                
                self.update_progress(25, "正在填写账号信息...")
                # 尝试多种用户名输入框
                user_selectors = [
                    'input[name="username"]',
                    'input[placeholder*="用户名"]',
                    'input[type="text"]'
                ]
                for sel in user_selectors:
                    try:
                        await page.fill(sel, self.config["username"], timeout=3000)
                        break
                    except:
                        continue
                
                # 填写密码
                await page.fill('input[type="password"]', self.config["password"])
                
                self.update_progress(35, "正在登录...")
                await page.click('button[type="submit"], button.el-button--primary, .login-btn')
                await asyncio.sleep(4)
                
                # 检查登录结果
                if "login" in page.url.lower() or await page.query_selector('input[type="password"]'):
                    raise Exception("登录失败，请检查账号密码或网络连接")
                
                self.update_progress(45, "登录成功！")
                
                # 2. 跳转到系统配置页
                self.update_progress(50, "正在跳转到入侵报警页面...")
                target_url = self.config["url"].rstrip('/') + "/portal/ui/system-configuration"
                await page.goto(target_url)
                await asyncio.sleep(3)
                
                # 3. 树状菜单导航
                self.update_progress(60, "正在展开设备管理...")
                await self.click_menu(page, "设备管理")
                await asyncio.sleep(1.5)
                
                self.update_progress(70, "正在点击报警检测...")
                await self.click_menu(page, "报警检测")
                await asyncio.sleep(1.5)
                
                self.update_progress(75, "正在点击入侵报警...")
                await self.click_menu(page, "入侵报警")
                await asyncio.sleep(5)  # 等待iframe/表格加载
                
                # 4. 查找并点击详情图标
                self.update_progress(85, "正在查找详情图标...")
                
                clicked = await self.click_detail_icon(page)
                if not clicked:
                    raise Exception("未找到详情图标（h-icon-details），请确认页面已加载完成")
                
                await asyncio.sleep(2)
                
                # 5. 点击在线获取
                self.update_progress(95, "正在点击在线获取...")
                clicked = await self.click_online_get(page)
                if not clicked:
                    raise Exception("未找到在线获取按钮")
                
                self.update_progress(100, "✅ 执行完成！")
                self.log("所有操作已成功执行")
                
                # 保持浏览器打开几秒让用户看到结果
                await asyncio.sleep(3)
                
                if browser:
                    await browser.close()
                    browser = None
                
        except Exception as e:
            self.log(f"❌ 错误: {str(e)}")
            if browser:
                try:
                    await browser.close()
                except:
                    pass
            raise
    
    async def click_menu(self, page, text):
        """点击菜单（支持多种选择器）"""
        selectors = [
            f'text="{text}"',
            f'//span[contains(text(),"{text}")]',
            f'//div[contains(text(),"{text}")]',
            f'//li[contains(text(),"{text}")]'
        ]
        
        for selector in selectors:
            try:
                if selector.startswith('//'):
                    await page.click(selector, timeout=3000)
                else:
                    await page.click(selector, timeout=3000)
                self.log(f"已点击菜单: {text}")
                return
            except:
                continue
        
        # 如果都失败，尝试JavaScript点击
        try:
            await page.evaluate(f'''
                () => {{
                    const items = document.querySelectorAll('*');
                    for (let item of items) {{
                        if (item.textContent.trim() === "{text}" || 
                            item.textContent.includes("{text}")) {{
                            item.click();
                            return true;
                        }}
                    }}
                }}
            ''')
            self.log(f"通过JS点击菜单: {text}")
        except Exception as e:
            raise Exception(f"无法点击菜单: {text}, {e}")
    
    async def click_detail_icon(self, page):
        """点击详情图标（支持iframe）"""
        # 获取所有frames（包括主页面）
        frames = [page] + page.frames[1:]  # 主页面 + 所有子frame
        
        for frame in frames:
            try:
                # 策略1：精确匹配 h-icon-details
                icons = await frame.query_selector_all('i.h-icon-details')
                if icons and len(icons) > 0:
                    await icons[0].click()
                    self.log("已点击 h-icon-details 图标")
                    return True
                
                # 策略2：模糊匹配
                icons = await frame.query_selector_all('i[class*="details"], i[class*="detail"]')
                if icons and len(icons) > 0:
                    await icons[0].click()
                    self.log("已点击详情图标（模糊匹配）")
                    return True
                
                # 策略3：表格第一行操作列
                rows = await frame.query_selector_all('table tbody tr, .el-table__row')
                if rows and len(rows) > 0:
                    # 获取第一行的所有单元格
                    cells = await rows[0].query_selector_all('td')
                    if cells and len(cells) > 0:
                        # 最后一列通常是操作列
                        last_cell = cells[-1]
                        # 在操作列内找图标或按钮
                        btn = await last_cell.query_selector('i, button, a')
                        if btn:
                            await btn.click()
                            self.log("已点击第一行操作按钮")
                            return True
                            
            except Exception as e:
                continue
        
        return False
    
    async def click_online_get(self, page):
        """点击在线获取按钮"""
        frames = [page] + page.frames[1:]
        
        for frame in frames:
            try:
                # 策略1：精确文本匹配
                btn = await frame.query_selector('button:has-text("在线获取")')
                if btn:
                    await btn.click()
                    self.log("已点击在线获取按钮")
                    return True
                
                # 策略2：所有按钮中查找
                buttons = await frame.query_selector_all('button, a')
                for btn in buttons:
                    text = await btn.text_content()
                    if text and "在线获取" in text:
                        await btn.click()
                        self.log("已点击在线获取（遍历查找）")
                        return True
                        
            except:
                continue
        
        return False
    
    def update_progress(self, value, message):
        """更新进度（线程安全）"""
        self.root.after(0, lambda: self._do_update(value, message))
    
    def _do_update(self, value, message):
        self.progress['value'] = value
        self.log(message)
    
    def reset_ui(self):
        """重置界面"""
        self.start_btn.config(state=tk.NORMAL, text="▶ 开始执行", bg="#1890ff")

if __name__ == "__main__":
    # 检查配置文件
    if not CONFIG_FILE.exists():
        default = {
            "url": "https://10.108.90.1",
            "username": "",
            "password": "",
            "headless": True,
            "slow_mo": 500
        }
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(default, f, ensure_ascii=False, indent=2)
        except:
            pass
    
    app = HikvisionApp()
    app.run()
