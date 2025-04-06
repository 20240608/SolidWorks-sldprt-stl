# SolidWorks-sldprt-stl
通过运行pyhon代码，可以实现在Windows系统安装SolidWorks软件的情况下，实现SLDPRT文件批量自动转换成STL文件，便于sw文件快速用于3D打印。


说明：

1. 您需要先配置python环境并安装pywin32库：pip install pywin32。


2. convert_sldprt_to_stl函数会遍历指定文件夹中的所有SLDPRT文件，并将它们转换为STL格式，保存到您输入的指定同一文件夹中。


3. 请确保SolidWorks已经安装，并且能够通过API访问。



运行步骤：

1. 命令行（cmd）输入：
   python sw_sldprt-stl.py
2.运行此脚本，会提示名称、最后一次编辑时间、运行环境、SolidWorks实例版本，然后通过与程序交互来实现文件夹路径输入。
3.SolidWorks将自动处理文件的转换,并保存到指定路径，同时记录本次的输入路径配置，自动生成配置文件，下次默认打开上一次的配置。
4.运行过程中会有进度提示和详细的报错日志。



此脚本适用于有SolidWorks环境的Windows系统。如果您没有SolidWorks许可或环境，可能需要使用其他转换工具。


