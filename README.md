# ClassGrade  
__由AI编写的一个Python班级个人成绩总结系统__   

### 项目结构
```
ClassGrade/
|── Grade/        --用来存放班级成绩Excel文件
├── 成绩模板.xlsx  --个人成绩模板
├── main.py       --主程序
├── README.md
```
### 运行环境
* Python 3.x
* pandas 库
* openpyxl 库
* xlrd 库

__安装依赖__  

Windows
```
pip install pandas openpyxl xlrd
```

Linux
```
python3 -m venv ClassGrade
source ClassGrade/bin/activate
pip install pandas openpyxl xlrd
```



### 注意事项
* 整个班级的成绩Excel文件必须不能有多余行，第一行为表头，例如：姓名、学号、成绩
* 目前只有部分年级和部分学科总结，如有需要可自行修改
* 程序会自动创建“学生成绩报告”文件夹，输出结果存放在里面