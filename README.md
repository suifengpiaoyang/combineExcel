# 合并 Excel 文件
A script to combine several Excel files for one.

要求:  
* 版本 : python3  
* 需求模块 : openpyxl  
* 后缀名：.xlsx

如果没有安装 openpyxl ，请先运行 `pip install openpyxl` 进行安装。  

使用方式：  
双击 combineExcel.py 首次运行将会生成一个 setting.py 文件。  

```python
# setting.py
FILE_SAVE_PATH = r''
NEED_COMBINE_DICTORY_PATH = r''
MODE = 1
```

在上面单引号之间填写路径。  
FILE_SAVE_PATH 对应着合并完成后文件的保存路径。如果不填，则默认保存在当前脚本的路径下。  
NEED_COMBINE_DICTORY_PATH 需要合并的文件所在的文件夹路径。如果不填，则默认目标路径为当前脚本所在路径。  
注意：   
* 生成的文件不会计算到需要合并的文件之中。  
* 能搜寻目标文件夹及其子文件夹下的 Excel 文件。  

MODE 表示模式选择，目前只有 0 和 1 两种模式：  
* 模式 0 : 普通合并所有 Excel 文件；
* 模式 1 : 合并所有 Excel 文件，从第二个文件开始不合并
         文件的第一行，因为有很多情况下这些 Excel 文
         件都是同样形式；

配置填写完成后，再次双击运行。无意外的话，就能得到合并后的文件。  
