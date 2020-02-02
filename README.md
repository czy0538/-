# 问卷星导出数据处理

最近新型肺炎爆发以后各地都要求进行一些关联史、流调之类的东西，而大部分科室还在采用最原始的逐个询问的方式，效率不仅低还增加了高危人群的接触风险，于是想到我们非常非常非常常用的问卷星，顺便利用python和vba实现了自动的处理，可以直接得到符合泰安卫健委要求的上报格式的文件。

其实所有的excle文件只需对程序稍加修改都可以使用这种方式进行批处理。程序的逻辑结构十分的简单，代码也没有什么难度，但可能对于我们父母这一代人来讲除了手动一个个处理外真的想不到这种方式~

## 零、运行环境要求

- Windows 10
- Office2010及更新版本

## 一、基本使用步骤

1. 在问卷星管理后台下载文件数据

   <img src="png\Annotation 2020-02-02 235532.png" style="zoom:50%;" />

   <img src="C:\Users\czy05\Desktop\问卷星格式化\png\Annotation 2020-02-02 235726.png" style="zoom:50%;" />



2. 将下载下来的文件放入“问卷星格式化”文件夹下

   <img src="C:\Users\czy05\Desktop\问卷星格式化\png\Annotation 2020-02-03 000002.png" style="zoom:50%;" />

   

3. 运行Formatting.exe，输入文件名并回车

   <img src="C:\Users\czy05\Desktop\问卷星格式化\png\Annotation 2020-02-03 000151.png" style="zoom:50%;" />

   ![](C:\Users\czy05\Desktop\问卷星格式化\png\Annotation 2020-02-03 000241.png)

4. 待窗口消失后，文件名.xlsx即为格式化后的文件

   <img src="C:\Users\czy05\Desktop\问卷星格式化\png\Annotation 2020-02-03 000355.png" style="zoom:50%;" />

## 二、自定义文件模版

1、建立一个示例原始文件，复制进“导出.xlsm”中

<img src="C:\Users\czy05\Desktop\问卷星格式化\png\Annotation 2020-02-03 001120.png" style="zoom:50%;" />

2、录制宏

<img src="C:\Users\czy05\Desktop\问卷星格式化\png\Annotation 2020-02-03 001224.png" style="zoom:50%;" />

3、 对宏进行修改并且命名为Formatting

<img src="C:\Users\czy05\Desktop\问卷星格式化\png\Annotation 2020-02-03 001338.png" style="zoom:50%;" />

4、 保存，运行Formatting即可
