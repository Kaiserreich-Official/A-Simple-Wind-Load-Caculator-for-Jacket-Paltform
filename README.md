## 这是一个简易的导管架平台风载荷计算器，可以批量计算多个构件的风载并将结果保存为Excel表格
# 注意
目前只能计算桩腿所受风载荷，尚未添加甲板和上部组块的计算功能
# 原理
将导管架平台的桩腿视为三维线段，利用罗德里格旋转矩阵根据风载荷方向进行投影，利用投影后的线段长度乘以管外径便得到受风面积
