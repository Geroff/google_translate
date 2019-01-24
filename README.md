>Python：Python3.7
>
>IDE：PyCharm Community版
>
需要request、execjs、xlrd、xlwt库
 
>pip install requests
>
>pip install PyExecJs
>
>pip install xlrd
>
>pip install xlwt
>
>如果没有pip命令，则需要将python根目录的Scripts目录添加到环境变量中

data:存储要翻译的Excel以及翻译后的Excel，带translate的为翻译的
excel_util.py：负责将需要翻译的内容从Excel读入，以及将结果存储到Excel
translate_google.py：负责翻译

备注：需要注意的是，有些字段在所有语言都是一样的，不需要翻译，因此需要最终还是要处理Excel，如果时间有限，需要做随机抽样检测，这样会保险
不适合频繁请求，会导致google超时，甚至IP被限制
