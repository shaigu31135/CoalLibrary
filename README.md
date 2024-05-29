# CoalLibrary使用手册



CoalLibrary（煤炭库）是一个轻量化的VB6动态链接库（COM组件）。

使用前，需要注册CoalLibrary.dll文件，并在VB的集成开发环境中添加引用。



### 一、字符串构造器类 StringBuilder

*注：此类只是简单地模仿了其它编程语言中字符串构造器的基本功能，其本质上还是字符串的类型。*



实例化：

```vb
Dim CL As New CoalLibrary.StringBuilder
```



方法和过程：

* Clear

	清除字符串构造器的全部内容。

	

* AppendText(text)

	将text追加至字符串构造器的末尾。

	```vb
	Dim CL As New CoalLibrary.StringBuilder
	CL.AppendText("123")
	CL.AppendText("456")
	```

	CL的内容将会是"123456"。

	

* Append(text)

	创建一个字符串构造器副本，然后将text追加至其末尾。

	你可以将若干个项拼接在一起：

	```vb
	Dim CL As New CoalLibrary.StringBuilder
	Dim s As New CoalLibrary.StringBuilder
	CL.AppendText("123")
	Set s = CL.Append("456").Append("789")
	```

	

* RecoverText(text)

	用text覆盖字符串构造器之前的内容。

	```vb
	Dim CL As New CoalLibrary.StringBuilder
	CL.AppendText("123")
	'此时CL的内容为"123"
	CL.RecoverText("456")
	'此时CL的内容为"456"
	```



* ToString

	获得字符串构造器的内容（返回值为字符串类型）。



* ReverseText

	反转字符串构造器内容的字符顺序。

```vb
Dim CL As New CoalLibrary.StringBuilder
CL.AppendText("123")
'此时CL的内容为"123"
CL.ReverseText()
'此时CL的内容为"321"
```



* Reverse

	创建一个字符串构造器副本，然后反转字符串构造器内容的字符顺序。

```vb
Dim CL As New CoalLibrary.StringBuilder
Dim s As New CoalLibrary.StringBuilder
CL.AppendText("123")
'此时CL的内容为"123"
Set s = CL.Reverse()
'此时s的内容为"321"
```



* SubString(start,length)

	获得字符串构造器内容的一部分，然后创建一个此部分的字符串构造器副本。

	参数length缺省时为-1，表示从start处取到字符串构造器内容末尾。

```vb
Dim CL As New CoalLibrary.StringBuilder
Dim s As New CoalLibrary.StringBuilder
CL.AppendText("12345")
Set s = CL.SubString(2,2)
'此时s的内容为"23"
s.Clear
Set s = CL.SubString(2)
'此时s的内容为"2345"
```



* UpCase/DownCase

	将字符串构造器内容的全部英文字母变成大/小写。



* FindText(text)/FindLastText(text)

	从左至右/从右至左查找遇到的第一个字符串text。如果没有查找到，返回-1。



* ReplaceText(find,text)

	用字符串text替换掉字符串构造器内容里面所有的字符串find。

```vb
Dim CL As New CoalLibrary.StringBuilder
CL.AppendText("12323")
'此时CL的内容为"12323"
CL.ReplaceText("2","6")
'此时CL的内容为"16363"
```



* ReplaceText(find,text)

	创建一个字符串构造器的副本，然后用字符串text替换掉这个副本的内容里面所有的字符串find。



属性：

* Length (只读)：字符串构造器内容的长度。



### 二、数学工具类 Mathematics

这是一个工具类，实例化后可以调用其中的函数、常量。

*注：圆周率π和自然对数底e为近似值。*



实例化：

```vb
Dim CL As New CoalLibrary.Mathematics
```



方法和过程：

* Logarithm(base,antilogarithm) 

	返回以base为底数，以antilogarithm为真数的对数。base是不等于1的正数。



* Exponential(base,power)

	返回base的power次方。



* RadToDeg(radian)和DegToRad(degree)

	弧度（角度）转换为角度（弧度）。



* SineD、CosineD、TangentD、SecantD、CosecantD、CotangentD和SineR、CosineR、TangentR、SecantR、CosecantR、CotangentR

	分别为正弦、余弦、正切、正割、余割、余切。后缀为D的函数表示输入的参数为角度制（Degree），后缀为R的函数表示输入的参数为弧度制（Radian）。



* HyperbolicSine、HyperbolicCosine、HyperbolicTangent、HyperbolicCotangent、HyperbolicSecant和HyperbolicCosecant

	双曲正弦、双曲余弦、双曲正切、双曲余切、双曲正割、双曲余割。



* Factorial(n)

	返回n!。



* Permutation(collection,sample)和Combination(collection,sample)

	排列和组合。collection为集合数，sample为样本数。

	```vb
	Dim CL As New CoalLibrary.Mathematics
	Msgbox CL.Permutation(5, 2)
	'将弹出消息框，内容为10
	```

	

* RandomInt(min,max)

	产生从min到max的伪随机整数。



* RandomSingle

	等价于vb自带的Rnd函数，但不是使用Randomize打乱伪随机数种子，而是使用当前时间作为伪随机数种子。



属性：

* PI

	圆周率的近似值（3.14159265358979）。



* E

	自然对数底的近似值（2.71828182845905）。

	

### 三、系统通用对话框SystemCommonDialog

*注：仅模仿了“打开文件”“保存文件”“颜色设置”和“字体设置”对话框，“打印设置”和“页面设置”因为不常用，所以并未添加。*

*此工具类源码参考自作者：Donald Grover，依赖系统的comdlg32.dll文件。*



* ShowOpenDialog(title,filter,initdir,mulselect,h_owner,h_instance)

	显示“打开文件”对话框。

	| 参数名称   | 描述                                | 备注                                 |
	| ---------- | ----------------------------------- | ------------------------------------ |
	| title      | 对话框标题，缺省为空                |                                      |
	| filter     | 文件过滤器，缺省为“所有文件\|\*.\*” | 和通用对话框控件的文件过滤器语法一致 |
	| initdir    | 初始化目录，缺省为空                |                                      |
	| mulselect  | 是否允许复选，缺省为False           |                                      |
	| h_owner    | 持有者窗体的句柄，缺省为0           | 通常为缺省值即可                     |
	| h_instance | 实例对象的句柄，缺省为0             | 通常为缺省值即可                     |

	返回选择的文件路径。如果mulselect设置为True，则返回选择的所有文件的路径，并以回车换行符（vbCrLf）分隔。



* ShowSaveDialog(title,filter,initdir,h_owner,h_instance)

​		显示“保存文件”对话框。参数描述参见ShowOpenDialog。

​		返回要保存的文件的路径。当目标文件存在时，将弹出覆盖询问提示框。如果单击“取消”按钮，返回空字符串。



* ShowColorDialog(h_owner,h_instance)

​		显示“颜色设置”对话框。参数描述参见ShowOpenDialog。

​		返回颜色值。如果单击“取消”按钮，返回-1。



* ShowFontDialog(h_owner)

	显示“字体设置”对话框。参数描述参见ShowOpenDialog。

	返回字体结构体（StdFont）。



### 四、字典Dictionary

类似于Python中的字典，但功能更简单。



### 五、字符串列表StringList

类似于列表框（ListBox）控件，但功能更简单。
