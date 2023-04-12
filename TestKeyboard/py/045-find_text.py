# 导入easyocr
import easyocr 
import sys
sys.setrecursionlimit(10000)




t1 = sys.argv[1] 
img_target = "{}".format(t1) 
# 创建reader对象
reader = easyocr.Reader(['ch_tra']) 
# 读取图像
# result = reader.readtext('./pic/mxd/test/sampleGuGu/1.jpg')
result = reader.readtext(img_target)
# 结果
print(result)
