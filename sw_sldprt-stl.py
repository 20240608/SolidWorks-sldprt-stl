import os
import sys
import traceback
import time
import json
import platform
from pathlib import Path

# 检查必要依赖并提供友好提示
try:
    import win32com.client
    import win32api
except ImportError:
    print("错误: 缺少必要的依赖库。")
    print("请安装以下库: pywin32")
    print("安装命令: pip install pywin32")
    sys.exit(1)

# 配置文件路径 - 使用pathlib提高跨平台兼容性
CONFIG_FILE = Path(__file__).parent / "sw_config.json"

# 显示程序信息和环境
script_path = Path(__file__).absolute()
script_name = script_path.name
script_name_without_ext = script_path.stem
last_modified_time = time.strftime("%H:%M:%S", time.localtime(os.path.getmtime(script_path)))

print(f"我是{script_name_without_ext}，最后一次编辑时间为：{last_modified_time}")
print(f"运行环境: Python {platform.python_version()} on {platform.system()} {platform.release()}")

# 加载配置函数
def load_config():
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
            print(f"已加载上次配置")
            return config
        except Exception as e:
            print(f"加载配置文件失败: {str(e)}")
    return {"input_directory": "", "output_directory": ""}

# 保存配置函数
def save_config(input_dir, output_dir):
    try:
        config = {
            "input_directory": input_dir,
            "output_directory": output_dir
        }
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
        print("已保存当前配置")
    except Exception as e:
        print(f"保存配置失败: {str(e)}")

def convert_sldprt_to_stl(input_directory, output_directory):
    swApp = None
    try:
        # 多级尝试连接SolidWorks
        connection_methods = [
            # 方法1: 连接到现有实例
            lambda: (win32com.client.GetActiveObject("SldWorks.Application"), "已连接到正在运行的SolidWorks实例"),
            # 方法2: 尝试创建指定版本(按新到旧顺序)
            lambda: (win32com.client.Dispatch("SldWorks.Application.31"), "已创建新的SolidWorks实例 (版本31)"),
            lambda: (win32com.client.Dispatch("SldWorks.Application.30"), "已创建新的SolidWorks实例 (版本30)"),
            lambda: (win32com.client.Dispatch("SldWorks.Application.29"), "已创建新的SolidWorks实例 (版本29)"),
            lambda: (win32com.client.Dispatch("SldWorks.Application.28"), "已创建新的SolidWorks实例 (版本28)"),
            lambda: (win32com.client.Dispatch("SldWorks.Application.27"), "已创建新的SolidWorks实例 (版本27)"),
            lambda: (win32com.client.Dispatch("SldWorks.Application.26"), "已创建新的SolidWorks实例 (版本26)"),
            # 方法3: 尝试创建默认版本
            lambda: (win32com.client.Dispatch("SldWorks.Application"), "已创建新的SolidWorks实例 (默认版本)")
        ]

        # 依次尝试各种连接方法
        for method in connection_methods:
            try:
                swApp, message = method()
                print(message)
                break
            except Exception:
                continue

        if swApp is None:
            print("无法连接到SolidWorks实例，请确保SolidWorks已正确安装。")
            return

        # 显示SolidWorks窗口
        swApp.Visible = True
        
        # 验证 SolidWorks 版本
        try:
            version_info = swApp.RevisionNumber
            print(f"SolidWorks版本: {version_info}")
        except:
            print("无法获取SolidWorks版本信息，但将继续尝试转换")
            
        # 使用Path对象处理路径，提高跨平台兼容性
        input_path = Path(input_directory)
        output_path = Path(output_directory)
            
        # 筛选有效的SLDPRT文件
        valid_files = []
        try:
            # 使用列表推导式更高效地筛选文件
            upper_files = [file for file in input_path.glob("*.SLDPRT") 
                          if not file.name.startswith("~$") and ".zip" not in file.name.lower()]
            lower_files = [file for file in input_path.glob("*.sldprt") 
                          if not file.name.startswith("~$") and ".zip" not in file.name.lower()]
            
            # 合并并去重
            valid_files = list(set(upper_files + lower_files))
        except Exception as e:
            print(f"扫描目录时出错: {str(e)}")
            traceback.print_exc()
            return
        
        if not valid_files:
            print(f"在目录 {input_directory} 中未找到有效的SLDPRT文件")
            return
        
        print(f"找到 {len(valid_files)} 个SLDPRT文件需要转换")
        total_files = len(valid_files)
        success_count = 0
        fail_count = 0
        start_time = time.time()

        for index, file_path in enumerate(valid_files, 1):
            # 计算进度和预计剩余时间
            try:
                percent_complete = (index - 1) / total_files * 100
                elapsed_time = time.time() - start_time
                if index > 1:
                    avg_time_per_file = elapsed_time / (index - 1)
                    est_time_remaining = avg_time_per_file * (total_files - (index - 1))
                    time_str = f"- 已用时: {elapsed_time:.1f}秒, 预计剩余: {est_time_remaining:.1f}秒"
                else:
                    time_str = ""
            except Exception as e:
                print(f"警告: 计算进度时出错: {str(e)}")
                time_str = ""
                
            file_name = file_path.name
            print(f"[{index}/{total_files} - {percent_complete:.1f}%] 处理文件: {file_name} {time_str}")
            
            # 转换处理开始
            try:
                # 这里使用WinAPI获取短路径（若失败则使用原始路径）
                try:
                    short_path = win32api.GetShortPathName(str(file_path))
                except Exception as e:
                    short_path = str(file_path)
                    print(f"注意: 无法获取短路径，使用原始路径。错误: {str(e)}")
                
                if not os.path.exists(short_path):
                    print(f"路径无效或文件不存在: {short_path}")
                    fail_count += 1
                    continue
                
                # 尝试关闭所有打开文档
                try:
                    swApp.CloseAllDocuments(True)
                except Exception as e:
                    print(f"警告: 无法关闭已打开的文档。错误: {str(e)}")
                    
                # 定义SolidWorks常量: swDocPART = 1
                swDocPART = 1
                
                # 尝试打开文档
                try:
                    # 方法1: 使用OpenDoc(更简单的接口)
                    model_doc = swApp.OpenDoc(short_path, swDocPART)
                    if model_doc is None:
                        try:
                            # 方法2: 使用不同参数组合的OpenDoc6
                            model_doc = swApp.OpenDoc6(short_path, swDocPART, 1, "", 0, 0)
                        except:
                            try:
                                # 方法3: 最简化参数的OpenDoc6
                                model_doc = swApp.OpenDoc6(short_path, swDocPART, 0, "", "", "")
                            except:
                                model_doc = None
                                
                    if model_doc is None:
                        print(f"无法打开文件: {file_path}")
                        fail_count += 1
                        continue
                except Exception as e:
                    print(f"尝试打开文件失败: {str(e)}")
                    fail_count += 1
                    continue
                
                # 定义输出STL文件路径
                stl_file_path = output_path / f"{file_path.stem}.STL"
                
                # 使用 SaveAs 方法进行导出
                ret = model_doc.SaveAs(str(stl_file_path))
                if ret:
                    success_count += 1
                    print(f"文件已成功转换为STL: {stl_file_path}")
                else:
                    fail_count += 1
                    print(f"文件转换失败: {file_path}")
            except Exception as e:
                print(f"处理文件 {file_name} 时出错: {str(e)}")
                traceback.print_exc()
                fail_count += 1
            finally:
                # 尝试关闭当前文档
                try:
                    swApp.CloseDoc(file_name)
                    print(f"已关闭文档: {file_name}")
                except Exception as e:
                    print(f"关闭文档时出错: {str(e)}")
                    
        print(f"批量转换完成! 成功: {success_count}, 失败: {fail_count}, 总计: {total_files}")
        
    except Exception as e:
        print(f"程序执行出错: {str(e)}")
        traceback.print_exc()
    finally:
        # 确保释放SolidWorks资源
        if swApp:
            try:
                swApp.ExitApp()  # 退出SolidWorks应用
                swApp = None
            except Exception as e:
                print(f"退出SolidWorks时出错: {str(e)}")

if __name__ == "__main__":
    try:
        # 加载上次的配置
        config = load_config()
        last_input_dir = config["input_directory"]
        last_output_dir = config["output_directory"]
        
        # 先获取输入文件夹路径
        if len(sys.argv) > 1:
            input_directory = sys.argv[1]
        else:
            # 显示上次的路径作为默认选项
            default_prompt = f" [按Enter使用默认值: {last_input_dir}]" if last_input_dir else ""
            input_prompt = f"请输入SLDPRT文件所在的文件夹路径{default_prompt}: "
            input_directory = input(input_prompt) or last_input_dir
        
        # 确认输入路径
        confirm_prompt = f"输入的SLDPRT文件所在目录是: {input_directory}，确认吗? (y/n) [默认y]: "
        confirm = input(confirm_prompt) or 'y'
        while confirm.lower() != 'y':
            default_prompt = f" [按Enter使用默认值: {last_input_dir}]" if last_input_dir else ""
            input_directory = input(f"请输入SLDPRT文件所在的文件夹路径{default_prompt}: ") or last_input_dir
            confirm = input(f"输入目录是: {input_directory}，确认吗? (y/n) [默认y]: ") or 'y'
        
        # 验证输入目录是否存在
        if not os.path.exists(input_directory):
            print(f"错误: 路径 {input_directory} 不存在!")
            sys.exit(1)
        
        # 让用户输入STL文件输出目录
        default_prompt = f" [按Enter使用默认值: {last_output_dir}]" if last_output_dir else ""
        output_directory = input(f"请输入STL文件输出的目录路径{default_prompt}: ") or last_output_dir
        confirm_out = input(f"输出目录是: {output_directory}，确认吗? (y/n) [默认y]: ") or 'y'
        while confirm_out.lower() != 'y':
            default_prompt = f" [按Enter使用默认值: {last_output_dir}]" if last_output_dir else ""
            output_directory = input(f"请输入STL文件输出的目录路径{default_prompt}: ") or last_output_dir
            confirm_out = input(f"输出目录是: {output_directory}，确认吗? (y/n) [默认y]: ") or 'y'
        
        # 验证输出目录是否存在，不存在则尝试创建
        if not os.path.exists(output_directory):
            try:
                os.makedirs(output_directory)
                print(f"输出目录 {output_directory} 已创建")
            except Exception as e:
                print(f"无法创建输出目录: {output_directory}，错误: {str(e)}")
                sys.exit(1)
        
        # 保存当前配置
        save_config(input_directory, output_directory)
        
        # 开始转换工作（先筛选有效文件，再转换）
        convert_sldprt_to_stl(input_directory, output_directory)
        
    except KeyboardInterrupt:
        print("\n程序被用户中断")
    except Exception as e:
        print(f"发生错误: {str(e)}")
        traceback.print_exc()
    
    input("按Enter键退出...")