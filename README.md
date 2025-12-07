# ClassGrade  
__由AI编写的一个Python班级个人成绩总结系统__   
支持总结班级个人成绩,总结前200名人数和人名  
目前支持的学期有初一-初三期中  

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
apt install python3.10-venv
python3 -m venv ClassGrade
source ClassGrade/bin/activate
pip install pandas openpyxl xlrd
```

### AI提示词  
若你想要使用AI进行修改,请将下面的提示词输入给ai  
```
这是一个使用python实现的班级个人成绩总结系统,可以实现通过班级成绩的Excel文件,总结班级个人成绩,总结前200名人数和人名,并将结果输出到Excel文件中。  
大体架构如下  
GradeSummaryGenerator类:用于实现班级个人成绩总结的功能,对于科目的定义在self.grade8_subjects,对于学期的定义在self.semester_keys  
GradeBefore200类:用于实现总结前200名人数和人名的功能,对于学期的定义在self.semesters  
main函数:用于程序的入口,实现对GradeSummaryGenerator类和GradeBefore200类的调用  

以下为我的修改要求  
(这里填充你的要求)  
```

### 注意事项
* 整个班级的成绩Excel文件必须不能有多余行，第一行为表头，例如：姓名、学号、成绩
* 目前只有部分年级和部分学科总结，如有需要可自行修改
* 程序会自动创建“学生成绩报告”文件夹，输出结果存放在里面