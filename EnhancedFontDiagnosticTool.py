import os
import subprocess
import platform
import shutil
import tempfile
import ctypes
import json
import hashlib
from collections import defaultdict
from pathlib import Path

class EnhancedFontDiagnosticTool:
    def __init__(self):
        self.system = platform.system()
        self.report = {
            'system': self.system,
            'issues': [],
            'suggestions': [],
            'missing_fonts': [],
            'corrupted_fonts': [],
            'font_cache_issues': False,
            'software_specific_issues': defaultdict(list),
            'dpi_scaling_issues': [],
            'font_integrity_issues': []
        }
        
        # 已知字体文件的MD5哈希值（示例值，实际应用中需要更完整的数据库）
        self.known_font_hashes = {
            "Arial.ttf": "a1b2c3d4e5f678901234567890123456",  # 示例哈希值
            "Times New Roman.ttf": "f6e5d4c3b2a198765432109876543210",
            "Segoe UI.ttf": "1234567890abcdef1234567890abcdef",
            "DejaVuSans.ttf": "abcdef1234567890abcdef1234567890",
            "FreeSans.ttf": "0987654321abcdef0987654321abcdef",
            "Helvetica.dfont": "aabbccddeeff00112233445566778899",
            "San Francisco.ttf": "11223344556677889900aabbccddeeff"
        }
        
    def run_full_diagnostics(self):
        """运行全面的字体诊断"""
        print("正在运行全面的字体诊断...")
        
        # 基本诊断
        self.check_font_directories()
        self.check_font_cache()
        self.check_system_fonts()
        self.check_font_config()
        self.check_locale_settings()
        
        # 扩展诊断
        self.check_software_specific_issues()
        self.check_font_integrity()
        self.check_dpi_scaling()
        
        print("\n诊断完成!")
        self.generate_report()
    
    def check_font_directories(self):
        """检查系统字体目录是否存在且可访问"""
        print("\n检查字体目录...")
        font_dirs = self.get_font_dirs()
        
        for directory in font_dirs:
            if not os.path.exists(directory):
                self.report['issues'].append(f"字体目录不存在: {directory}")
                self.report['suggestions'].append(f"创建目录: {directory} 或修复系统字体设置")
            elif not os.access(directory, os.R_OK):
                self.report['issues'].append(f"无法访问字体目录: {directory}")
                self.report['suggestions'].append(f"检查目录权限: {directory}")
            else:
                print(f"字体目录正常: {directory}")
    
    def get_font_dirs(self):
        """获取系统字体目录路径"""
        if self.system == "Windows":
            return [
                os.path.join(os.environ['WINDIR'], 'Fonts'),
                os.path.join(os.environ['LOCALAPPDATA'], 'Microsoft', 'Windows', 'Fonts')
            ]
        elif self.system == "Linux":
            return [
                '/usr/share/fonts',
                '/usr/local/share/fonts',
                os.path.expanduser('~/.fonts')
            ]
        elif self.system == "Darwin":  # macOS
            return [
                '/Library/Fonts',
                '/System/Library/Fonts',
                os.path.expanduser('~/Library/Fonts')
            ]
        else:
            return []
    
    def check_font_cache(self):
        """检查字体缓存问题"""
        print("\n检查字体缓存...")
        
        if self.system == "Windows":
            # Windows字体缓存可以通过重建来修复
            cache_dir = os.path.join(os.environ['LOCALAPPDATA'], 'Microsoft', 'Windows', 'FontCache')
            if not os.path.exists(cache_dir):
                self.report['font_cache_issues'] = True
                self.report['issues'].append("Windows字体缓存目录不存在")
            else:
                print("Windows字体缓存目录存在")
        
        elif self.system == "Linux":
            # 检查常见的Linux字体缓存
            try:
                subprocess.run(['fc-cache', '-v'], check=True, capture_output=True, text=True)
                print("Linux字体缓存正常")
            except subprocess.CalledProcessError as e:
                self.report['font_cache_issues'] = True
                self.report['issues'].append("Linux字体缓存问题")
                print(f"字体缓存错误: {e.stderr}")
        
        elif self.system == "Darwin":
            # macOS通常不需要手动管理字体缓存
            print("macOS字体缓存通常自动管理")
    
    def check_system_fonts(self):
        """检查基本系统字体是否存在"""
        print("\n检查系统字体...")
        
        required_fonts = {
            "Windows": ["Arial.ttf", "Times New Roman.ttf", "Segoe UI.ttf"],
            "Linux": ["DejaVuSans.ttf", "FreeSans.ttf"],
            "Darwin": ["Helvetica.dfont", "San Francisco.ttf"]
        }.get(self.system, [])
        
        missing_fonts = []
        
        for font in required_fonts:
            found = False
            for font_dir in self.get_font_dirs():
                if os.path.exists(os.path.join(font_dir, font)):
                    found = True
                    break
            
            if not found:
                missing_fonts.append(font)
        
        if missing_fonts:
            self.report['missing_fonts'] = missing_fonts
            self.report['issues'].append(f"缺少关键系统字体: {', '.join(missing_fonts)}")
            self.report['suggestions'].append("考虑重新安装系统或恢复默认字体")
        else:
            print("关键系统字体存在")
    
    def check_font_config(self):
        """检查系统字体配置"""
        print("\n检查字体配置...")
        
        if self.system == "Linux":
            try:
                # 检查fontconfig配置
                result = subprocess.run(['fc-match', 'sans-serif'], capture_output=True, text=True)
                if result.returncode != 0:
                    self.report['issues'].append("字体配置问题: 无法匹配基本字体")
                    print("字体配置问题:", result.stderr)
                else:
                    print("基本字体匹配正常:", result.stdout.strip())
            except FileNotFoundError:
                self.report['issues'].append("fontconfig未安装或不可用")
                self.report['suggestions'].append("安装fontconfig包: sudo apt install fontconfig")
    
    def check_locale_settings(self):
        """检查区域设置是否可能影响字体显示"""
        print("\n检查区域设置...")
        
        if self.system == "Linux":
            try:
                result = subprocess.run(['locale'], capture_output=True, text=True)
                if "en_US.UTF-8" not in result.stdout:
                    self.report['issues'].append("区域设置可能不支持某些字体")
                    self.report['suggestions'].append("考虑设置LANG=en_US.UTF-8")
            except Exception as e:
                print("无法检查区域设置:", str(e))
    
    def generate_report(self):
        """生成诊断报告"""
        print("\n=== 诊断报告 ===")
        print(f"系统: {self.report['system']}")
        
        if self.report['issues']:
            print("\n发现的问题:")
            for issue in self.report['issues']:
                print(f"- {issue}")
        else:
            print("\n未发现重大问题")
        
        if self.report['missing_fonts']:
            print("\n缺失的字体:")
            for font in self.report['missing_fonts']:
                print(f"- {font}")
        
        if self.report['corrupted_fonts']:
            print("\n可能损坏的字体:")
            for font in self.report['corrupted_fonts']:
                print(f"- {font}")
        
        if self.report['font_integrity_issues']:
            print("\n字体完整性问题:")
            for font in self.report['font_integrity_issues']:
                print(f"- {font} 可能已损坏")
        
        if self.report['dpi_scaling_issues']:
            print("\nDPI/缩放问题:")
            for issue in self.report['dpi_scaling_issues']:
                print(f"- {issue}")
        
        if self.report['software_specific_issues']:
            print("\n软件特定问题:")
            for software, issues in self.report['software_specific_issues'].items():
                print(f"{software}:")
                for issue in issues:
                    print(f" - {issue}")
        
        if self.report['suggestions']:
            print("\n建议的解决方案:")
            for suggestion in self.report['suggestions']:
                print(f"- {suggestion}")
        
        # 将报告保存到文件
        report_path = os.path.join(os.path.expanduser("~"), "font_diagnostic_report.txt")
        with open(report_path, 'w') as f:
            json.dump(self.report, f, indent=2)
        print(f"\n完整报告已保存到: {report_path}")
    
    def fix_font_cache(self):
        """尝试修复字体缓存"""
        print("\n尝试修复字体缓存...")
        
        if self.system == "Windows":
            print("在Windows上重建字体缓存:")
            print("1. 打开命令提示符(管理员)")
            print("2. 运行: net stop FontCache")
            print("3. 运行: del /q %windir%\\System32\\FNTCACHE.DAT")
            print("4. 运行: net start FontCache")
            print("5. 重启计算机")
        
        elif self.system == "Linux":
            try:
                print("在Linux上重建字体缓存...")
                subprocess.run(['fc-cache', '-f', '-v'], check=True)
                print("字体缓存重建成功")
            except subprocess.CalledProcessError as e:
                print("重建字体缓存失败:", e.stderr)
        
        elif self.system == "Darwin":
            print("在macOS上重建字体缓存:")
            print("1. 打开终端")
            print("2. 运行: atsutil databases -remove")
            print("3. 重启计算机")
    
    def install_font(self, font_path):
        """安装新字体"""
        print(f"\n尝试安装字体: {font_path}")
        
        if not os.path.exists(font_path):
            print("错误: 字体文件不存在")
            return
        
        font_dirs = self.get_font_dirs()
        if not font_dirs:
            print("错误: 无法确定系统字体目录")
            return
        
        target_dir = font_dirs[0]  # 使用第一个字体目录
        font_name = os.path.basename(font_path)
        target_path = os.path.join(target_dir, font_name)
        
        try:
            if not os.path.exists(target_dir):
                os.makedirs(target_dir)
            
            shutil.copy2(font_path, target_path)
            print(f"字体已安装到: {target_path}")
            
            # 更新字体缓存
            if self.system == "Linux":
                subprocess.run(['fc-cache', '-f', '-v'], check=True)
            
            self.report['suggestions'].append(f"新字体安装完成: {font_name}")
        except Exception as e:
            print(f"安装字体失败: {str(e)}")
            self.report['issues'].append(f"字体安装失败: {font_name}")
    
    def restore_default_fonts(self):
        """恢复默认系统字体"""
        print("\n恢复默认系统字体...")
        
        if self.system == "Windows":
            print("在Windows上恢复默认字体:")
            print("1. 打开设置 > 个性化 > 字体")
            print("2. 点击'恢复默认字体设置'")
            print("3. 或者从另一台相同系统的计算机复制字体")
        
        elif self.system == "Linux":
            print("在Linux上恢复默认字体:")
            print("1. 重新安装字体包:")
            print("   Debian/Ubuntu: sudo apt install --reinstall fonts-dejavu fonts-freefont-ttf")
            print("   Fedora/RHEL: sudo dnf reinstall dejavu-sans-fonts freefont")
        
        elif self.system == "Darwin":
            print("在macOS上恢复默认字体:")
            print("1. 从Time Machine备份恢复/Library/Fonts和/System/Library/Fonts")
            print("2. 或重新安装操作系统")
    
    def check_software_specific_issues(self):
        """检查特定软件的字体问题"""
        print("\n检查特定软件的字体问题...")
        
        # Adobe 软件检查
        if self.system == "Windows":
            if os.path.exists(r"C:\Program Files\Adobe"):
                print("检测到Adobe软件安装")
                self.check_adobe_fonts()
        
        # Microsoft Office 检查
        if self.system == "Windows":
            if os.path.exists(r"C:\Program Files\Microsoft Office"):
                print("检测到Microsoft Office安装")
                self.check_office_fonts()
        
        # 开发工具检查
        self.check_ide_fonts()
    
    def check_adobe_fonts(self):
        """检查Adobe软件的字体问题"""
        try:
            # 检查Adobe字体目录
            adobe_font_dir = os.path.join(os.environ.get('APPDATA', ''), 'Adobe', 'CoreSync', 'plugins', 'livetype', 'c')
            if os.path.exists(adobe_font_dir):
                print(f"检查Adobe字体目录: {adobe_font_dir}")
                
                # 检查是否有损坏的字体缓存
                cache_files = [f for f in os.listdir(adobe_font_dir) if f.endswith('.lst')]
                if not cache_files:
                    self.report['software_specific_issues']['Adobe'].append("Adobe字体缓存文件缺失")
                    self.report['suggestions'].append("尝试在Adobe软件中重置字体首选项")
        except Exception as e:
            print(f"检查Adobe字体时出错: {str(e)}")
    
    def check_office_fonts(self):
        """检查Microsoft Office的字体问题"""
        try:
            # 检查Office使用的字体是否可用
            office_fonts = ["Calibri.ttf", "Cambria.ttf", "Consolas.ttf"]
            missing = []
            
            for font in office_fonts:
                found = False
                for font_dir in self.get_font_dirs():
                    if os.path.exists(os.path.join(font_dir, font)):
                        found = True
                        break
                
                if not found:
                    missing.append(font)
            
            if missing:
                self.report['software_specific_issues']['Microsoft Office'].append(
                    f"Office使用的字体缺失: {', '.join(missing)}")
                self.report['suggestions'].append("重新安装Microsoft Office或手动安装缺失的Office字体")
        except Exception as e:
            print(f"检查Office字体时出错: {str(e)}")
    
    def check_ide_fonts(self):
        """检查开发工具(IDE)的字体问题"""
        common_ides = {
            "VSCode": {
                "paths": [
                    os.path.expanduser("~/.vscode"),
                    os.path.expanduser("~/AppData/Roaming/Code")
                ],
                "config_file": "settings.json",
                "font_setting": "editor.fontFamily"
            },
            "IntelliJ": {
                "paths": [
                    os.path.expanduser("~/.IntelliJIdea*"),
                    os.path.expanduser("~/Library/Preferences/IntelliJIdea*")
                ],
                "config_file": "options/editor.xml",
                "font_setting": "FONT_FAMILY"
            }
        }
        
        for ide, info in common_ides.items():
            for path_pattern in info["paths"]:
                # 处理通配符路径
                if '*' in path_pattern:
                    import glob
                    matches = glob.glob(path_pattern)
                    if not matches:
                        continue
                    ide_path = matches[0]
                else:
                    ide_path = path_pattern
                
                config_path = os.path.join(ide_path, info["config_file"])
                if os.path.exists(config_path):
                    print(f"检测到{ide}安装，检查字体配置...")
                    self.check_ide_font_config(ide, config_path, info["font_setting"])
                    break
    
    def check_ide_font_config(self, ide_name, config_path, font_setting_key):
        """检查IDE的字体配置"""
        try:
            font_family = ""
            
            if config_path.endswith('.json'):
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    font_family = config.get(font_setting_key, "")
            elif config_path.endswith('.xml'):
                # 使用更健壮的XML解析
                try:
                    import xml.etree.ElementTree as ET
                    tree = ET.parse(config_path)
                    root = tree.getroot()
                    
                    for elem in root.iter(font_setting_key):
                        font_family = elem.text.strip()
                        break
                except ImportError:
                    # 回退到简单文本搜索
                    with open(config_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                        import re
                        match = re.search(f"{font_setting_key}[^>]*>([^<]+)", content)
                        font_family = match.group(1) if match else ""
            
            if font_family:
                # 检查配置的字体是否可用
                available = False
                for font_dir in self.get_font_dirs():
                    for ext in ['.ttf', '.otf', '.ttc']:
                        if os.path.exists(os.path.join(font_dir, f"{font_family}{ext}")):
                            available = True
                            break
                    if available:
                        break
                
                if not available:
                    self.report['software_specific_issues'][ide_name].append(
                        f"配置的字体 '{font_family}' 不可用")
                    self.report['suggestions'].append(
                        f"在{ide_name}设置中更改为已安装的字体")
        except Exception as e:
            print(f"检查{ide_name}字体配置时出错: {str(e)}")
    
    def check_font_integrity(self):
        """检查字体文件的完整性"""
        print("\n检查字体文件完整性...")
        
        # 只检查系统关键字体
        critical_fonts = {
            "Windows": ["Arial.ttf", "Times New Roman.ttf", "Segoe UI.ttf"],
            "Linux": ["DejaVuSans.ttf", "FreeSans.ttf"],
            "Darwin": ["Helvetica.dfont", "San Francisco.ttf"]
        }.get(self.system, [])
        
        for font in critical_fonts:
            for font_dir in self.get_font_dirs():
                font_path = os.path.join(font_dir, font)
                if os.path.exists(font_path):
                    if not self.verify_font_integrity(font_path, font):
                        self.report['font_integrity_issues'].append(font)
                    break
    
    def verify_font_integrity(self, font_path, font_name):
        """验证字体文件的完整性"""
        try:
            # 计算文件的MD5哈希
            with open(font_path, 'rb') as f:
                file_hash = hashlib.md5(f.read()).hexdigest()
            
            # 与已知好的哈希值比较
            known_hash = self.known_font_hashes.get(font_name)
            if known_hash and file_hash != known_hash:
                print(f"⚠ 字体文件可能损坏: {font_path}")
                print(f"  期望哈希: {known_hash}")
                print(f"  实际哈希: {file_hash}")
                return False
            
            return True
        except Exception as e:
            print(f"验证字体完整性时出错 ({font_path}): {str(e)}")
            return False
    
    def check_dpi_scaling(self):
        """检查DPI和显示缩放设置"""
        print("\n检查DPI和显示缩放设置...")
        
        if self.system == "Windows":
            try:
                import winreg
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, 
                                   "Control Panel\\Desktop") as key:
                    dpi_value = winreg.QueryValueEx(key, "LogPixels")[0]
                    if dpi_value not in [96, 120, 144]:  # 常见标准DPI值
                        self.report['dpi_scaling_issues'].append(
                            f"非标准DPI设置 ({dpi_value}) 可能导致字体显示问题")
                        self.report['suggestions'].append(
                            "尝试将DPI设置为100%(96)、125%(120)或150%(144)")
            except Exception as e:
                print(f"无法检查Windows DPI设置: {str(e)}")
        
        elif self.system == "Linux":
            try:
                # 尝试获取GNOME的缩放设置
                result = subprocess.run(['gsettings', 'get', 'org.gnome.desktop.interface', 'scaling-factor'], 
                                       capture_output=True, text=True)
                if result.returncode == 0:
                    scaling = result.stdout.strip()
                    if scaling not in ["1", "2"]:  # 常见缩放值
                        self.report['dpi_scaling_issues'].append(
                            f"非标准缩放设置 ({scaling}) 可能导致字体显示问题")
            except Exception as e:
                print(f"无法检查Linux缩放设置: {str(e)}")
        
        elif self.system == "Darwin":
            try:
                # 获取macOS的显示缩放设置
                result = subprocess.run(['system_profiler', 'SPDisplaysDataType'], 
                                      capture_output=True, text=True)
                if "Resolution" in result.stdout:
                    import re
                    match = re.search(r"Resolution:\s*(.+)", result.stdout)
                    if match:
                        resolution = match.group(1)
                        if "scaled" in resolution.lower():
                            self.report['dpi_scaling_issues'].append(
                                "检测到缩放的显示分辨率，可能影响字体显示")
            except Exception as e:
                print(f"无法检查macOS显示设置: {str(e)}")
    
    def fix_dpi_scaling(self):
        """尝试修复DPI/缩放设置问题"""
        print("\n尝试修复DPI/缩放设置...")
        
        if not self.report['dpi_scaling_issues']:
            print("未发现DPI/缩放设置问题")
            return
        
        if self.system == "Windows":
            print("Windows DPI/缩放设置修复:")
            print("1. 右键桌面选择'显示设置'")
            print("2. 在'缩放与布局'部分")
            print("3. 将'更改文本、应用等项目的大小'设置为推荐值(通常是100%或125%)")
            print("4. 点击'高级缩放设置'")
            print("5. 输入自定义缩放值(如果需要)，然后注销并重新登录")
        
        elif self.system == "Linux":
            print("Linux 缩放设置修复:")
            print("GNOME桌面:")
            print("1. 打开设置 > 设备 > 显示")
            print("2. 调整缩放比例为1或2")
            print("3. 或者使用命令: gsettings set org.gnome.desktop.interface scaling-factor 1")
        
        elif self.system == "Darwin":
            print("macOS 显示缩放修复:")
            print("1. 打开系统偏好设置 > 显示器")
            print("2. 选择'默认'或'更多空间'选项")
            print("3. 避免使用缩放的选项")

def main():
    print("=== 增强版字体问题诊断与修复工具 ===")
    tool = EnhancedFontDiagnosticTool()
    
    while True:
        print("\n菜单:")
        print("1. 运行全面诊断")
        print("2. 修复字体缓存")
        print("3. 安装新字体")
        print("4. 恢复默认字体")
        print("5. 修复DPI/缩放设置")
        print("6. 检查特定软件字体问题")
        print("7. 退出")
        
        choice = input("请选择操作 (1-7): ")
        
        if choice == '1':
            tool.run_full_diagnostics()
        elif choice == '2':
            tool.fix_font_cache()
        elif choice == '3':
            font_path = input("输入字体文件路径: ").strip()
            tool.install_font(font_path)
        elif choice == '4':
            tool.restore_default_fonts()
        elif choice == '5':
            tool.fix_dpi_scaling()
        elif choice == '6':
            # 显示已检测到的软件问题
            if tool.report['software_specific_issues']:
                print("\n检测到的软件特定问题:")
                for software, issues in tool.report['software_specific_issues'].items():
                    print(f"{software}:")
                    for issue in issues:
                        print(f" - {issue}")
            else:
                print("未检测到特定软件的字体问题")
        elif choice == '7':
            print("退出程序")
            break
        else:
            print("无效选择，请重试")

if __name__ == "__main__":
    main()