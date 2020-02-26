



# 微信公众号

------

微信公众号：VBA168

淘宝店铺地址：https://item.taobao.com/item.htm?spm=a1z10.1-c-s.w4004-21233576391.4.1af0683dzrx3oU&id=584940166162

关注微信公众号，每天及时接收Excel VBA经典示例讲解。

淘宝店铺提供Excel定制服务。

祝你工作和学习更轻松！

------



# Excel对象

我们每天重复着打开、关闭工作簿，输入、清除单元格内容的操作，其实都是在操作Excel的对象。如下操作：

- 退出Excel程序是操作Application对象；
- 新建工作簿是操作Workbook对象；
- 新建工作表是操作Worksheet对象；
- 删除单元格是操作Range对象；
- 插入图表是操作Chart对象。

实际上，VBA程序就是用代码记录下来的一个或一串操作，如想在"Sheet1"工作表的A1单元格输入数值"100"，完整的代码为：

`Application.Worksheets("Sheet1").Range("A1").Value=100`

无论是用动作还是代码完成任务，都是在操作对象。所以编写VBA程序，就是用VBA语句引用对象并有目的地操作它。

我们日常工作中经常会操作下表所列的对象。

| 对象        | 对象说明                                                     |
| ----------- | ------------------------------------------------------------ |
| Application | 代表Excel应用程序(如果在Word中使用VBA，就代表Word应用程序)   |
| Workbook    | 代表Excel的工作簿，一个Workbook对象代表一个工作簿文件        |
| Worksheet   | 代表Excel的工作表，一个Worksheet对象代表工作簿中的一张普通工作表 |
| Range       | 代表Excel中的单元格，可以是单个单元格，也可以是单元格区域。  |

记住这些常用的对象，并能熟练地运用代码操作它们，就可以执行一些常用的操作。

## Application对象

Application对象代表Excel程序本身，它就像一棵树的根，Excel中所有的对象都以它为起点。实际编程时，会经常用到它的许多属性和方法。

### 1.用ScreenUpdating属性设置是否更新屏幕上的内容

在使用Excel解决一个问题时，往往需要执行多步操作或计算。无论是通过手动还是VBA代码完成这些操作，默认情况下，Excel都会将每步操作所得的结果显示在屏幕上。

Application对象的ScreenUpdating属性就是控制屏幕更新的开关，将该属性设置为False，Excel将会关闭屏幕更新。

```
Application.ScreenUpdating=False   '关闭屏幕更新

Application.ScreenUpdating=True    '重新开启屏幕更新
```

如果在程序中将ScreenUpdating属性设置为False，一定要记得在程序结束前将其重新设置为True，否则后面执行的程序也不会更新屏幕上的内容。

### 2.设置DisplayAlerts属性禁止显示警告对话框

当我们在Excel中执行某些操作，如删除工作表时，Excel会显示一个警告对话框，让我们确定是否需要执行这个操作，如图1所示。

![1](D:\微信公众号\别怕，Excel VBA起始很简单\4.2\1.JPG)

图1 删除工作表时显示的警告对话框

出于很多原因，我们都希望Excel在程序执行的过程中，不显示警告对话框，这可以通过设置Application对象的DisplayAlerts属性为False来实现。

在下面示例代码中，如果不写第2行代码，那么在执行程序中，每删除一个工作表，就会弹出一次警告框，而且只有点击【删除】按钮才会执行删除操作。如果要删除100张工作表，就需要点击【删除】按钮100次。

```vb
Sub DelSht_2()
    Application.DisplayAlerts = False           '设置不显示警告对话框
    Dim sht As Worksheet
    For Each sht In Worksheets
        If sht.Name <> ActiveSheet.Name Then  '判断sht引用的是否是活动工作表
            sht.Delete                        '删除sht引用的工作表
        End If
    Next sht
    Application.DisplayAlerts = True            '重新设置显示警告对话框
End Sub
```

如果在程序中将DisplayAlerts属性设置为False，一定要记得在程序结束前将其重新设置为True，否则后面执行任何操作都不会弹出警告对话框。

### 3.借助WorksheetFunction属性使用工作表函数

常用的工作表函数，如SUMIF、VLOOKUP、MATCH和COUNTIF等函数VBA中没有，但是可以使用Application对象的WorksheetFunction属性来调用这些函数。



图2 编写代码统计A1:B50单元格区域中大于1000的数据个数

如想统计图2中A1:B50单元格区域中大于1000的数值有多少个，代码可以写为：

```vb
Sub CountTest_2()
    Dim mycount As Integer
    mycount = Application.WorksheetFunction.CountIf(Range("A1:B50"), ">1000")
    MsgBox "A1:B50中大于1000的单元格个数为：" & mycount
End Sub
```

注：如果VBA中已经有了相同功能的函数，就不能再通过WorksheetFunction属性引用工作表中的函数，否则会出错。并且，不是所有的工作表函数都能通过WorkshetFunction 属性来调用。

## Workbook对象

Workbooks是所有工作簿对象组成的集合，而Wrokbook对象是Workbooks集合的一个成员。

### 1.引用集合中的工作簿

利用工作簿名引用工作簿，如已经打开了"Book1.xlsm"工作簿，那么`Workbooks("Book1.xlsm")`就代表这个工作簿对象。

### 2.访问对象的属性，获得工作簿文件的信息

通过代码获得指定工作簿的名称、保存的路径等文件信息，示例代码如下。

```vb
Sub WbMsg()
    Range("B2") = ThisWorkbook.Name '获得工作簿的名称
    Range("B3") = ThisWorkbook.Path '获得工作簿文件所在的路径
    Range("B4") = ThisWorkbook.FullName '获得带路径的工作簿名称
End Sub
```

代码中的ThisWorkbook代表代码所在的工作簿对象。

注：后缀为.xlsx的工作簿无法保存VBA代码。如果工作簿中编写了宏代码，则将工作簿保存成后缀为.xls或者.xlsm的文件。

### 3.用Add方法创建工作簿

```vb
Workbooks.Add '创建空白工作簿
```

```vb
Workbooks.Add "D:\模板.xlsm" '指定用来创建工作簿的模板
```

### 4.用Open方法打开工作簿

打开一个Excel工作簿文件，最简单的方法是使用Workbooks对象的Open方法，示例代码入戏。

```vb
Workbooks.Open Filename:="D:\模板.xlsm"
```

方法Open和参数Filename之间用空格分隔，参数及参数数值之间用":="连接。

在实际使用时，代码中的参数名称Filename可以省略不写，将代码写为：

```vb
Workbooks.Open "D:\模板.xlsm"
```

更改代码中的路径及文件名称，即可打开其他的工作簿文件。

### 5.用Activate方法激活工作簿

虽然可以同时打开多个工作簿，但同一时间只能有一个工作簿是活动的。如果想让不活动的工作簿变为活动工作簿，可以用Wrokbooks对象的Activate方法激活它。例如：

```vb
Workbooks("模板.xlsm").Activate
```

### 6.保存工作簿文件

用Save方法保存已经存在的文件，例如：

```vb
ThisWorkbook.Save '保存代码所在的工作簿
```

用SaveAs方法将工作簿另存为新文件。如果是第1次保存一个新建的工作簿，或需要将工作簿另存为一个新文件时，应该使用SaveAs方法，例如：

```vb
ThisWorkbook.SaveAs Filename:="D:\Test.xlsm"
```

另存新文件后不关闭原文件。使用SaveAs方法将工作簿另存为新文件后，Excel将关闭原文件并自动打开另存为得到的新文件，如果希望继续保留原文件不打开新文件，应该使用**SaveCopyAs**方法。例如：

```vb
ThisWorkbook.SaveCopyAs Filename:="D:\Test.xlsm"
```

### 7.用Close方法关闭工作簿

调用工作簿对象的Close方法，可以关闭打开的工作簿。例如：

```vb
Workbooks.Close '关闭当前打开的所有工作簿
```

可以通过索引号、名称等指定要打开的工作簿，例如：

```vb
Workbooks("Test.xlsm").Close '关闭名称为Test的工作簿
```

```vb
Workbooks("Test.xlsm").Close savechange:=True '关闭并保存对工作簿的修改
```

### 8. ThisWorkbook与ActiveWorkbook

两者都返回Workbook对象。不同之处在于，ThisWorkbook是对**代码所在工作簿**的引用，ActiveWorkbook是对**活动工作簿**的引用。

# 实例代码

## #1.使用 QRMaker.ocx控件生成二维码

```
Sub qrCodeTest()


  Dim QRstr As String


  Const fgfStr As String = "
&
^
&
"


  QRstr = "A9508" + fgfStr + "董建华" + fgfStr + "朝阳区酒仙桥4街坊14楼2单元39、40" + fgfStr + "100016"


  With Application.Workbooks("appTestQrmaker1.xls").Worksheets("Sheet1")


  .QRmaker1.ModelNo = 2


  .QRmaker1.CellPitch = 5


  .QRmaker1.CellUnit = 203


  .QRmaker1.QuietZone = 0


​


  .QRmaker1.AutoRedraw = True


  .QRmaker1.InputData = QRstr


  .QRmaker1.Refresh


  .QRmaker1.CreateQrMetaFile 1, Application.ActiveWorkbook.Path 
&
 "\" 
&
 "A9508.bmp", 2


  End With


End Sub
```

## 2.裁剪图片

```
ActiveSheet.Cells(1, 1).Select


  ActiveSheet.Pictures.Insert( _


  "D:\My Documents\My Pictures\6cf0492712220dc1f03f5b.jpg").Select




  With Selection.ShapeRange




 '裁剪


  .PictureFormat.CropTop = 30  '下移裁剪


  .PictureFormat.CropLeft = 30 '右移裁剪


  .PictureFormat.CropBottom = 30 '上移裁剪


  .PictureFormat.CropRight = 30 '左移裁剪


 '裁剪




  '移动旋转  通常移动距离都是和裁剪相对应的，这样图才能在指定单元格的位置。


  .IncrementLeft -30  '相对图片初始位置水平移动正数向右，负数向左


  .IncrementTop -30 '相对图片初始位置垂直移动正数向下，负数向上


  .IncrementRotation 0 '相对图片初始位置中心旋转


  '移动旋转




  '大小


  .LockAspectRatio = msoFalse '图片纵横比锁定为msoTrue,高度和宽度调一个值整个图就会变


  .Height = 200 ' 高度


  .Width = 150  '宽度


  '大小


 End With
```

```
ActiveSheet.Cells(3, 1).Select  '选择要插入图片的单元格，定位


i = "D:\My Documents\My Pictures\111.gif" '图片地址可以写入变量


ActiveSheet.Pictures.Insert(i).Select 
 '用变量插图片


Cells(1, 1) = Selection.Name 
  '在第一个单元格返回插入图片对象的名称，方便以后的操作


Application.CommandBars("Picture").Visible = False 
 '隐藏图片编辑工具


ActiveSheet.Shapes(Cells(1, 1)).Select 
 '选择第一个单元格里所留的那个名称的图片对象


Selection.Delete 
  '删除选择的图片对象


删除全部图片的一种方法


Dim Sh As Shape 
 '定义一个图形的变量


For Each Sh In ActiveSheet.Shapes 
 '遍游活动表里的所有图形组件


  If Sh.Name Like "Picture *" Then  '如果图形对象的名称里有“Picture *”通配的往下执行，因为图片对象默认对象名称是Picture 数字


  Sh.Select 
  '选择图片名称的对象


  Selection.Delete 
 '删除图片对象


  End If


Next Sh 
  '利用循环就把图片对象都给删除了。
```

## 3.生成票号

```
Public Function getNewNum(numType As String, yuanNum As String) As String
    Dim dangRi As String
    dangRi = CStr(Format(Date, "yymmdd"))
    Dim xinNum As String
    xinNum = "001"
    Dim qianStr As String
    Dim isChaXun As Integer
    Dim jiNum As String
    Dim jiRi As String   

    qianStr = Trim(CStr(ThisWorkbook.Sheets("金额转换页").Range("B5").Value))
    jiNum = Trim(CStr(ThisWorkbook.Sheets("金额转换页").Range("B1").Value))
    isChaXun = Val(ThisWorkbook.Sheets("金额转换页").Range("C1").Value)
    jiRi = Trim(CStr(ThisWorkbook.Sheets("金额转换页").Range("B3").Value))
    If dangRi = jiRi Then
        If (Right(yuanNum, 3) <> jiNum) And (isChaXun < 1) Then
            xinNum = CStr(Format(Val(jiNum) + 1, "000"))
        Else
            xinNum = jiNum
        End If
    End If
    ThisWorkbook.Sheets("金额转换页").Range("B1").Value = xinNum
    ThisWorkbook.Sheets("金额转换页").Range("B3").Value = dangRi

    getNewNum = qianStr & dangRi & xinNum

End Function
```

```
Sub 生成票号()
    Dim RiQiStr As String
    RiQiStr = Trim(Range("B3"))
    Range("B1") = getNewNum("", RiQiStr)
End Sub
```

## 将多张工作表中的数据合并到一张工作表中

在工作中，我们经常遇到多张工作表合并到一张工作表的问题，比如希望将下图所示中各分表中保存的成绩记录，汇总到工作簿中的"成绩表"工作表中，可以用以下程序。

```vb
Sub hebing()
    '把各班成绩表中的记录合并到"成绩表"工作表中
    Dim sht As Worksheet
    Set sht = Worksheets("成绩表")
    sht.Rows("2:" & sht.rows.count).Clear      '删除成绩表中的原有记录
    Dim wt As Worksheet, xrow As Integer, rng As Range
    For Each wt In Worksheets                   '循环处理工作簿中的每张工作表
        If wt.Name <> "成绩表" Then
            Set rng = sht.Range("A1048576").End(xlUp).Offset(1, 0)
            xrow = wt.Range("A1").CurrentRegion.Rows.Count - 1
            wt.Range("A2").Resize(xrow, 7).Copy rng
        End If
    Next
End Sub
```

第4行代码意思是将"成绩表"工作表赋值给sht对象，在VBA中，给对象赋值，前面必须加**Set**关键字。

第5行代码中的`sht.rows.count`表示sht工作表总共有多少行；在VBA中，Rows表示工作表或某个区域中所有行组成的集合。`Rows("2:3")`表示选中工作表的第2行到第3行。

第7行代码中的wt代表工作表集合中的一个工作表，随着循环变换。

第9行代码表示wt工作表中数据A列下面的第一个空单元格。`Range("A1048576")`表示工作表最下面一个单元格。

Range对象的End属性返回包含指定单元格的区域最尾端的单元格，返回结果等同于在单元格中按【End+方向键】(上、下、左、右方向键)组合键得到的单元格。

Range对象的Offset属性获得相对于安远隔区域一定偏移位置上的单元格区域。`Offset(1, 0)`表示单元格下面一个单元格。

第10行表示A1单元格所在当前区域的行数键1。

第11行表示将子表中的数据复制到汇总表的空白区域。

Range对象的Resize属性将指定的单元格区域有目的地扩大或缩小，得到一个新的单元格区域。`Range("B2").Resize(5,4).Select`表示将B2单元格扩展为一个5行4列的单元格区域。



# 通过文件对话框选择文件

```vb
Sub 选择文件()
    Application.DisplayAlerts = False
    Dim fd As FileDialog, fn As String
    Dim wb As Workbook, sht As Worksheet, n As Long
    Dim i As Long, IRow As Long
    Dim wt As Worksheet
    Dim strFind As String, strReplace As String
    Dim strSchool As String
    Dim JRow As Long, j As Long
    Dim arr() As Variant
    Dim s1 As String
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "打开文件"
        .AllowMultiSelect = False
        .InitialFileName = ThisWorkbook.Path & "\"
        .Filters.Clear
        .Filters.Add "Excel文件", "*.xls;*.xlsx"
    End With
    If fd.Show = -1 Then
        For n = 1 To fd.SelectedItems.Count
            fn = fd.SelectedItems(n)
            Set wb = Workbooks.Open(fn)
            Set sht = wb.Worksheets(1)
            Set wt = ThisWorkbook.Worksheets(1)
            IRow = wt.Range("A" & wt.Rows.Count).End(xlUp).Row
            ********************

            ********************
            wb.Save
            'wb.Close
        Next
        Beep
        MsgBox "已完成！"
    Else
        MsgBox "请选择文件"
        Exit Sub
    End If
    Application.DisplayAlerts = True
End Sub
```

## 限制工作簿的使用次数

大多数试用版的软件通常都有使用时间或使用次数的限制

```
新增工作簿的文档属性


​


Sub AddCustomDocumentProperties()


  ThisWorkbook.CustomDocumentProperties.Add _


  Name:="OpenTimes", _


  LinkToContent:=False, _


  Type:=msoPropertyTypeNumber, _ '出现未定义变量，改成1


  Value:=0


End Sub


​


2.工作簿打开时触发以下事件


Private Sub Workbook_Open()


  Dim intOpentimes As Integer


  With Me


  intOpentimes = .CustomDocumentProperties _


  ("Opentimes").Value + 1


  If intOpentimes 
>
 3 Then'大于3时，执行工作簿“自杀代码”


  .Saved = True


  .ChangeFileAccess xlReadOnly


  Kill .FullName'删除磁盘上的工作簿文件


  .Close False


  Else


  .CustomDocumentProperties("Opentimes"). _'保存工作簿打开次数


  Value = intOpentimes


  .Save


  End If


  End With


End Sub


​
```

## 字典

```
Dim d As Object
Set d = CreateObject("Scripting.Dictionary")
```

### Exitsts方法

## 数组

将单元格区域赋值给数组进行处理，会大幅提高计算速度。

```
Dim arr() As Variant
arr = Range("A1:A4")
```

## 使用InputBox函数显示对话框供用户输入数据

```
Sub SimpleInput()


  Dim strInput As String


  strInput = InputBox("请输入邮政编码:", "邮政编码", "100001")


  If StrPtr(strInput) 
<
>
 0 Then


  If IsNumeric(Trim(strInput)) = True Then


  If Len(Trim(strInput)) = 6 Then


  Cells(1, 1) = strInput


  Else


  MsgBox "邮政编码必须是6位!"


  End If


  Else


  MsgBox "邮政编码必须是数值!"


  End If


  Else


  MsgBox "你已放弃了输入!"


  End If


End Sub


​
```

## Application.Interactive属性，禁止所有的键盘输入和鼠标操作

```
Application.Interactive=False'使应用程序处于非交互模式


Application.Interactive=True'宏代码运行结束不会自动恢复为True，需要在过程结束前将该属性值重新设置为True。


​


​
```

## run方法

使用Call方法调用过程时，被调用的过程名称不能使用变量。使用Run方法运行变量指定的过程。`Application.Run "Macro" & intIndex`

## OnKey方法捕捉键盘输入

利用OnKey方法捕捉键盘输入，能够实现当用户在Excel中按特定键或组合键时运行指定的过程。'其中参数Key可以指定与&lt;Alt&gt;、&lt;Ctrl&gt;或&lt;Shift&gt;组合使用的键，每一个键可由一个或多个字符表示，如“a”表示字符a,或者“{ENTER}”表示&lt;Enter&gt;键捕捉&lt;Ctrl+V&gt;组合键："^v"

`Application.OnKey (Key, [Procedure])`

## SendKeys方法模拟键盘输入

语法如下参数String为字符串表达式，用来指定要发送的按键消息。`Application.SendKeys(String, [Wait]) Application.SendKeys "%las%V~" '表示模拟键盘一次输入<ALter+l><a><s><Alt+V>和<Enter>`

## 巧妙捕获用户中断

在代码运行期间，如果用户按&lt;Esc&gt;键或&lt;Ctrl+Break&gt;组合键，即可中断代码执行

## 保存指定工作表到新的工作簿文件

```
  ActiveSheet.Copy


  ActiveWorkbook.Close SaveChanges:=True, _


  Filename:=ThisWorkbook.Path 
&
 "\SheetBackup.xlsx"


End Sub


​
```

## 查看指定的工作簿是否存在

```
Function blnExistWB(ByVal strWbName As String) As Boolean


  Dim wkbName As Workbook


  On Error Resume Next


  Set wkbName = Workbooks.Open(strWbName)


  If Err.Number = 0 Then blnExistWB = True


  Set wkbName = Nothing


End Function
```

## 打开工作簿时禁止更新链接

```vb
Application.AskToUpdateLinks=False'关闭更新数据源链接的提示
Application.AskToUpdateLinks=True'将属性值恢复为默认状态
```

# 单元格插入超链接

注：如果工作簿和工作表名称中包含空格等特殊字符，则还需要在外侧加单引号。

```vb
   Sub 添加超链接()
    Dim i As Long
    Dim ShtName As String
    Dim Sht As Worksheet
    Dim wt As Worksheet
    Set wt = Sheet1
    For i = 2 To 30
        ShtName = Trim(wt.Cells(i, 1))
        If Not SheetExists(ShtName) Then
            ThisWorkbook.Worksheets.Add after:=Worksheets(Worksheets.Count)
            ActiveSheet.Name = ShtName

        End If
        Set Sht = Worksheets(ShtName)
        wt.Activate
        wt.Hyperlinks.Add Anchor:=wt.Cells(i, 1), Address:="", SubAddress:=Sht.Name & "!A1", TextToDisplay:=wt.Cells(i, 1).Value
        Sht.Hyperlinks.Add Anchor:=Sht.Range("A1"), Address:="", SubAddress:=wt.Name & "!A1", TextToDisplay:="返回"
    Next
End Sub
```

## Split函数与数组的结合使用

```vb
Dim FileName As String, Arr As Variant
Arr = Split(FileName, "-")
PrefixStr = Arr(0)
```

## DisplayAlerts属性设置不显示警告框

```vb
Application.DisplayAlerts = False
```



# 第4章 Range操作

## 4.2取得最后一个非空单元格

xlDown/xlToRight/xlToLeft/xlUp

```vb
Dim ERow as Long
Erow=Range("A" & Rows.Count).End(xlUp).Row
```

## 4.3 复制单元格区域

注意：使用PasteSpecial方法时指定xlPasteAll（粘贴全部），并不包括粘贴列宽

```vb
Sub CopyWithSameColumnWidths()
    Sheets("Sheet1").Range("A1").CurrentRegion.Copy
    With Sheets("Sheet2").Range("A1")
        .PasteSpecial xlPasteColumnWidths
        .PasteSpecial xlPasteAll
    End With
    Application.CutCopyMode = False
End Sub
Sheets("Sheet2").Range("A1").PasteSpecial Paste:=xlPasteValues '粘贴数值
```

## 4.9 设置字符格式

### 4.9.1设置单元格文本字符串格式

```vb
Sub CellCharacter()
    With Range("A1")
        .Clear
        .Value = "Y=X2+1"
        .Characters(4, 1).Font.Superscript = True '将第4个字符设置为上标
        .Characters(1, 1).Font.ColorIndex = 3
        .Font.Size = 20
    End With
End Sub
```

通过Range对象的Characters属性来操作指定的字符。

Characters属性返回一个Characters对象，代表对象文字的字符区域。Characters属性的语法格式如下

```
Characters(Start, Length)
```

### 4.9.2 设置图形对象文本字符格式

如下示例为A3单元格批注添加指定文本，并设置字符格式。

```vb
Sub ShapeCharacter()
    If Range("A3").Comment Is Nothing Then
        Range("A3").AddComment Text:=""
    End If
    With Range("A3").Comment
        .Text Text:="Microsoft Excel 2016"
        .Shape.TextFrame.Characters(17).Font.ColorIndex = 3'返回从第17个字符开始到最后一个字符的字符串
    End With
End Sub
```

TextFrame属性返回Shape对象的文本框对象，而Characters属性返回其中的文本字符。

## 4.10 单元格区域添加边框

使用Range对象的Borders集合可以快速地对单元格区域全部边框应用相同的格式。

Range对象的BorderAround方法可以快速地为单元格区域添加外边框。

```
Sub AddBorders()
    Dim rngCell As Range
    Set rngCell = Range("B2:F8")
    With rngCell.Borders
        .LineStyle = xlContinuous '边框线条的样式
        .Weight = xlThin '设置边框线条粗细
        .ColorIndex = 5 '设置边框线条颜色
    End With
    rngCell.BorderAround xlContinuous, xlMedium, 5 '添加一个加粗外边框
    Set rngCell = Nothing
End Sub
```

![image-20200206164323610](C:\Users\admin\AppData\Roaming\Typora\typora-user-images\image-20200206164323610.png)

在单元格区域中应用多种边框格式

```vb
Sub BordersIndexDemo()
    Dim rngCell As Range
    Set rngCell = Range("B2:F8")
    With rngCell.Borders(xlInsideHorizontal) '内部水平
        .LineStyle = xlDot
        .Weight = xlThin
        .ColorIndex = 5
    End With
    With rngCell.Borders(xlInsideVertical) '内部垂直
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 5
    End With
    rngCell.BorderAround xlContinuous, xlMedium, 5
    Set rngCell = Nothing
End Sub
```

Borders\(index\)属性返回单个Border对象，其参数index取值可为以下：

| 名称 | 值 | 说明 |
| :--- | :--- | :--- |
| **xlDiagonalDown** | 5 | 从区域中每个单元格的左上角到右下角的边框。 |
| **xlDiagonalUp** | 6 | 从区域中每个单元格的左下角到右上角的边框。 |
| **xlEdgeBottom** | 9 | 区域底部的边框。 |
| **xlEdgeLeft** | 7 | 区域左边缘的边框。 |
| **xlEdgeRight** | 10 | 区域右边缘的边框。 |
| **xlEdgeTop** | 8 | 区域顶部的边框。 |
| **xlInsideHorizontal** | 12 | 区域中所有单元格的水平边框（区域以外的边框除外）。 |
| **xlInsideVertical** | 11 | 区域中所有单元格的垂直边框（区域以外的边框除外）。 |

去除边框

```vb
Sub Restore()
    Columns("B:F").Borders.LineStyle = xlNone
End Sub
```

## 4.11 高亮显示单元格区域

高亮显示是指以某种方式突出显示活动单元格或指定的单元格区域，使得用户可以一目了然地获取某些信息。

1.高亮显示单个单元格

```
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Cells.Interior.ColorIndex = xlNone'清除所有单元格的内部填充颜色
    Target.Interior.ColorIndex = 5
End Sub
```

![image-20200206165636905](C:\Users\admin\AppData\Roaming\Typora\typora-user-images\image-20200206165636905.png)

2.高亮显示行列

```
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim rngHighLight As Range
    Dim rngCell1 As Range, rngCell2 As Range
    Cells.Interior.ColorIndex = xlNone
    Set rngCell1 = Intersect(ActiveCell.EntireColumn, _
        [HighLightArea])
    Set rngCell2 = Intersect(ActiveCell.EntireRow, [HighLightArea])
    On Error Resume Next
    Set rngHighLight = Application.Union(rngCell1, rngCell2)
    rngHighLight.Interior.ThemeColor = 9
    Set rngCell1 = Nothing
    Set rngCell2 = Nothing
    Set rngHighLight = Nothing
End Sub
```

命名区域HighLightArea\(示例文件已指定B2:H15单元格区域\)

![image-20200206165756300](C:\Users\admin\AppData\Roaming\Typora\typora-user-images\image-20200206165756300.png)

3.结合条件格式定义名称高亮显示行

```
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ThisWorkbook.Names.Add "ActRow", ActiveCell.Row
End Sub
```

![image-20200206165917049](C:\Users\admin\AppData\Roaming\Typora\typora-user-images\image-20200206165917049.png)

4.结合条件格式定义名称高亮显示行列

```
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ThisWorkbook.Names.Add "ActRow", ActiveCell.Row
    ThisWorkbook.Names.Add "ActCol", ActiveCell.Column
End Sub
```

![image-20200206170134713](C:\Users\admin\AppData\Roaming\Typora\typora-user-images\image-20200206170134713.png)

## 4.12 动态设置单元格数据验证序列

【数据验证】对话框如下图

![image-20200206171335869](C:\Users\admin\AppData\Roaming\Typora\typora-user-images\image-20200206171335869.png)

如下示例代码通过VBA将示例工作簿中工作表“Office 2016"以外的工作表名称设置为工作表“Office 2016"中C3单元格的数据验证序列。

数据验证序列是由逗号分隔的字符串，两个逗号之间的空字符串将被忽略。

```vb
Sub SheetsNameValidation()
    Dim i As Integer
    Dim strList As String
    Dim wksSht As Worksheet
    For Each wksSht In Worksheets
        If wksSht.Name <> "Office 2016" Then
            strList = strList & wksSht.Name & ","
        End If
    Next wksSht
    With Worksheets("Office 2016").Range("C3").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:=strList
    End With
    Set wksSht = Nothing
End Sub
```

```
Sub DeleteValidation()
    Range("C3").Validation.Delete
End Sub
```

![image-20200206171703131](C:\Users\admin\AppData\Roaming\Typora\typora-user-images\image-20200206171703131.png)

Validation对象的Add方法向指定区域内添加数据验证，其语法格式如下：

```
Add (Type, AlertStyle, Operator, Formula1, Formula2)
```

参数Type是必需的，代表数据验证类型。其值可为以下常量之一：

| 名称 | 值 | 说明 |
| :--- | :--- | :--- |
| **xlValidateCustom** | 7 | 使用任意公式验证数据有效性。 |
| **xlValidateDate** | 4 | 日期值。 |
| **xlValidateDecimal** | 2 | 数值。 |
| **xlValidateInputOnly** | 0 | 仅在用户更改值时进行验证。 |
| **xlValidateList** | 3 | 值必须存在于指定列表中。 |
| **xlValidateTextLength** | 6 | 文本长度。 |
| **xlValidateTime** | 5 | 时间值。 |
| **xlValidateWholeNumber** | 1 | 全部数值。 |

参数Formula2指定数据验证公式的第二部分。仅当Operator为xlBetween或xlNotBetween时有效。

## 4.14 判断单元格公式是否存在错误

Excel公式返回的结果可能是一个错误的文本，包含\#NULL、\#DIV/0!、\#VALUE！、\#REF！、\#NAME?、\#NUM！和\#N/A等。

通过判断Range对象中的Value属性的返回结果是否为错误值，可得知公式是否存在错误。

```VB
Sub FormulaIsError()
    If VBA.IsError(Range("A1").Value) = True Then
        MsgBox "A1单元格错误类型为:" & Range("A1").Text
    Else
        MsgBox "A1单元格公式结果为:" & Range("A1").Value
    End If
End Sub
```

IsError函数判断表达式是否为一个错误值，如果是则返回逻辑值True，否则返回逻辑值False。

## 4.15批量删除所有错误值

使用CurrentRegion属性取得包含A1单元格的当前区域。

```
Sub DeleteError()
    Dim rngRange As Range
    Dim rngCell As Range
    Set rngRange = Range("a1").CurrentRegion
    For Each rngCell In rngRange
        If VBA.IsError(rngCell.Value) = True Then
            rngCell.Value = ""
        End If
    Next rngCell
    Set rngCell = Nothing
    Set rngRange = Nothing
End Sub
```

通过定位功能可获取错误值的单元格对象，并批量修改。

利用单元格对象的SpecialCells方法定位所有错误值。

```vb
Sub DeleteAllError()
    On Error Resume Next
    Dim rngRange As Range
    Set rngRange = Range("a1").CurrentRegion.SpecialCells _
        (xlCellTypeConstants, xlErrors)
    If Not rngRange Is Nothing Then
        rngRange.Value = ""
    End If
    Set rngRange = Nothing
End Sub
```

单元格对象的SpecialCells方法返回一个Range对象，该对象代表与指定类型和值匹配的所有单元格，其语法格式如下：

```vb
SpecialCells(Type,Value)
```

参数与Type是必需的，用于指定定位类型，可为如下表列举的XlCellType常量之一。

| 常量                           | 值    | 说明                       |
| ------------------------------ | ----- | -------------------------- |
| xlCellTypeAllFormatConditions  | -4172 | 任何格式的单元格           |
| xlCellTypeAllValidation        | -4174 | 含有验证条件的单元格       |
| xlCellTypeBlanks               | 4     | 空单元格                   |
| xlCellTypeComments             | -4144 | 含有注释的单元格           |
| xlCellTypeConstants            | 2     | 含有常量的单元格           |
| xlCellTypeFormulas             | -4123 | 含有公式的单元格           |
| xlCellTypeLastCell             | 11    | 已用区域中的最后一个单元格 |
| xlCellTypeSameFormatConditions | -4173 | 具有相同的格式的单元格     |
| xlCellTypeSameValidation       | -4175 | 验证条件相同的单元格       |
| xlCellTypeVisible              | 12    | 所有可见单元格             |

如果参数Type为xlCellTypeConstants或xlCellTypeFormulas，则该参数可用于确定结果中应包含哪几类单元格，参数Value可为以下列举的XlSpecialCellsValue常量之一。将这些值相加可使此方法返回多种类型的单元格。默认情况下，将选择所有常量或公式，无论类型如何。

| 常量             | 值   | 说明                 |
| :--------------- | :--- | :------------------- |
| **xlErrors**     | 16   | 有错误的单元格。     |
| **xlLogical**    | 4    | 具有逻辑值的单元格。 |
| **xlNumbers**    | 1    | 具有数值的单元格。   |
| **xlTextValues** | 2    | 具有文本的单元格。   |

## 4.17 判断单元格是否存在批注

```vb
Function blnComment(ByVal rngRange As Range) As Boolean
    If rngRange.Cells(1).Comment Is Nothing Then
        blnComment = False
    Else
        blnComment = True
    End If
End Function
```

返回单元格区域rngRange的第一个单元格是否存在批注。

注：对于合并单元格的批注，批注对象从属于合并单元格的第一个单元格。

Range对象的Comment属性返回批注对象，如果指定的单元格不存在批注，则该属性返回Nothing。

## 4.18 为单元格添加批注

```vb
Sub Comment_Add()
    With Range("B5")
        If .Comment Is Nothing Then
            .AddComment Text:=.Text
            .Comment.Visible = True
        End If
    End With
End Sub
```

使用Range对象的AddComment方法为单元格添加批注。

## 编辑批注文本

使用批注对象的Text方法，能够获取或修改单元格批注的文本。

```vb
Sub Comment_Add()
    With Range("B5")
        If .Comment Is Nothing Then
            .AddComment Text:=.Text
            .Comment.Visible = True
        End If
    End With
End Sub
```

Comment对象的Text方法的语法格式如下。

```
Text(Text,Start,Overwrite)
```

参数Text代表需要添加的文本。

参数Start指定添加文本的起始位置。

参数OrverWrite指定是否覆盖现有文本。默认值为False(新文字插入现有文字中)。

vbCrLf常量代表回车换行符。

## 4.21 显示图片批注

为单元格批注添加背景图片或将图片作为批注的内容

```vb
Sub ChangeCommentShapeType()
    With Range("B3").Comment
        .Shape.Fill.UserPicture _
            ThisWorkbook.Path & "\Logo.jpg"
    End With
End Sub
```

Comment对象的Shape属性返回批注对象的图形对象

Fill属性能够返回FillFormat对象，该对象包括指定的图表或图形的填充格式属性，UserPicture方法为图形填充图像

## 4.22 设置批注字体

单元格批注的字体通过单元格批注的Shape对象中文本框对象(TextFrame)的字符对象(Characters)进行设置。TextFrame代表Shape对象中的文本框，包含文本框中的文字。

```vb
Sub CommentFont()
    Dim objComment As Comment
    For Each objComment In ActiveSheet.Comments
        With objComment.Shape.TextFrame.Characters.Font
            .Name = "微软雅黑"
            .Bold = msoFalse
            .Size = 14
            .ColorIndex = 3
        End With
    Next objComment
    Set objComment = Nothing
End Sub

```

## 4.23 快速判断单元格区域是否存在合并单元格

Range对象的MergeCells属性可以判断单元格区域是否包含合并单元格，如果该属性返回值为True，则表示区域包含合并单元格。

```vb
Sub IsMergeCell()
    If Range("A1").MergeCells = True Then
        MsgBox "包含合并单元格"
    Else
        MsgBox "没有包含合并单元格"
    End If
End Sub
```

对于单个单元格，直接通过MergeCells属性判断是否包含合并单元格。

```vb
Sub IsMerge()
    If VBA.IsNull(Range("A1:E10").MergeCells) = True Then
        MsgBox "包含合并单元格"
    Else
        MsgBox "没有包含合并单元格"
    End If
End Sub
```

当单元格区域中同时包含合并单元格和非合并单元格时，MergeCells属性将返回Null.

## 4.24合并单元格时连接每个单元格内容

在合并多个单元格时，将各个单元格的内容连接起来保存在合并后的单元格区域中。

```vb
Sub MergeValue()
    Dim strText As String
    Dim rngCell As Range
    If TypeName(Selection) = "Range" Then
        For Each rngCell In Selection
            strText = strText & rngCell.Value
        Next rngCell
        Application.DisplayAlerts = False
        Selection.Merge
        Selection.Value = strText
        Application.DisplayAlerts = True
    End If
    Set rngCell = Nothing
End Sub
```

使用TypeName函数判断当前选定对象是否为Range对象。

将DisplayAlerts属性设置为False，禁止Excel弹出警告对话框。

## 4.25 取消合并时在每个单元格中保留内容

```vb
Sub UnMergeValue()
    Dim strText As String
    Dim i As Long, intCount As Integer
    For i = 2 To Range("B1").End(xlDown).Row
        With Cells(i, 1)
            strText = .Value
            intCount = .MergeArea.Count
            .UnMerge
            .Resize(intCount, 1).Value = strText
        End With
        i = i + intCount - 1
    Next i
End Sub
```

## 4.26 合并内容相同的单列连续单元格

```vb
Sub BackUp()
    Dim intRow As Integer, i As Long
    Application.DisplayAlerts = False
    With ActiveSheet
        intRow = .Range("A1").End(xlDown).Row
        For i = intRow To 2 Step -1
            If .Cells(i, 1).Value = .Cells(i - 1, 1).Value Then
                .Range(.Cells(i - 1, 1), .Cells(i, 1)).Merge
            End If
        Next i
    End With
    Application.DisplayAlerts = True
End Sub
```

使用For循环结构从最后一行开始，向上逐个判断相邻单元格内容的内容是否相同，如果相同则合并单元格区域。

## 4.27 查找包含指定字符串的所有单元格

查找指定单元格区域中所有包含“BB"的单元格并通过背景颜色标识。使用Range对象的Find方法可以实现此要求。

```vb
Sub FindCells()
    Dim rngCell As Range, rngResult As Range
    Dim strFirstAddress As String
    With Range("A1:E10")
        Set rngCell = .Find(What:="BB", After:=.Cells(1), _
            LookIn:=xlValues, LookAt:=xlPart)
        If Not rngCell Is Nothing Then
            strFirstAddress = rngCell.Address'保存第一个匹配单元格的引用地址。
            Set rngResult = rngCell
            Do
                Set rngResult = Application.Union(rngResult, rngCell)
                Set rngCell = .FindNext(rngCell)
            Loop While rngCell.Address <> strFirstAddress'设置当查找到的单元格与第一个匹配单元格地址相同时停止查找。
            rngResult.Interior.ColorIndex = 8'为单元格指定单元格内部填充颜色。
        End If
    End With
    Set rngCell = Nothing
    Set rngResult = Nothing
End Sub
```

Range对象的Find方法可以在单元格区域中查找特定的信息，并返回Range对象。如果未发现匹配单元格，则返回Nothing。

代码中的FindNext方法进行由Find方法设置的搜索，查找匹配相同条件的下一个单元格。当FindNext方法查找到指定查找区域的末尾时，该方法将重新从区域的第一个单元格继续搜索。

应用于Range对象的Find方法的语法格式如下。

```
Find(What, After, LookIn, LookAt, SearchOrder, SearchDirection, MatchCase, MatchByte, SearchFormat)
```

其中，参数What指定要搜索的数据，可为字符串或任意 Microsoft Excel 数据类型。

参数After用于确定开始搜索的位置，搜索过程将从该参数指定的单元格之后进行。参数After必须是区域中的单个单元格，从该单元格之后开始搜索，直到Find方法绕回到指定的单元格时，才结束搜索。

参数LookIn指定要查找的信息类型，可为如下列举的xlFindLookin常量之一。

| 名称           | 值    | 说明   |
| -------------- | ----- | ------ |
| **xlComments** | -4144 | 批注。 |
| **xlFormulas** | -4123 | 公式。 |
| **xlValues**   | -4163 | 值。   |

参数LookAt可为xlWhole(完整匹配)或xlPart(部分匹配)。

参数SearchOrder可为xlByRows(按行)或xlByColumns(按列)。

参数SearchDirection指定搜索的方向，可为xlNext(向后，默认值)或xlPrevious(向前)。

参数MatchCase指定查找时是否区分大小写，若为True则区分大小写，默认值为False。

参数MatchByte指定是否进行字节匹配。若为True，则双字节字符仅匹配双字节字符。若为  **False**，则双字节字符可与其对等的单字节字符匹配。

单字节指只占一个字，是英文字符。双字是占两个字节的，中文字符都占两个字节。

参数SearchFormat指定搜索格式

注：每次使用Find方法后，参数LookIn、LookAt、SearchOrder和MatchByte的设置将被保存。

## 4.30 按指定条件自动筛选数据

当需要在数据列表中检索符合指定条件的记录时，通常可以使用自动筛选功能实现。在vba中，使用Range集合的AutoFilter方法，可以实现对下图所示的数据区域执行自动筛选操作。

如下示例代码筛选"性别"字段内容为"男"的记录。

```vb
Sub DataFilter()
    Worksheets("Sheet1").Range("A1").AutoFilter _
        Field:=2, Criteria1:="男", Operator:=xlFilterValues
End Sub
```

当AutoFilter方法应用于单个单元格时，Excel默认筛选区域为该单元格所在的当前区域，参数Field指定作为筛选基准字段(从列表左侧开始，最左侧的字段为第一个字段)的偏移量，此处参数值为2，指定筛选字段在筛选区域的第二列。参数Criteria1指定筛选条件为"男"。

Range对象的AutoFilter方法筛选出符合条件的数据列表，其语法格式如下。

```
AutoFilter(Field, Criteria1, Operator, Criteria2, VisibleDropDown)
```

参数Field是可选的，为Variant类型，用于指定作为筛选基准字段的偏移量。

参数Criteria1是可选的，为Variant类型，用于指定筛选条件(字符串，"男")。使用"="可搜索到空字段，使用"<>"可搜索到非空字段。如果省略该参数，则搜索条件为All。

参数Operator是可选的，其值可为如下表列举的XlAutoFilterOperator常量之一。

| 名称                  | 值   | 说明                                    |
| --------------------- | ---- | --------------------------------------- |
| **xlAnd**             | 1    | 条件 1 和条件 2 的逻辑与。              |
| **xlBottom10Items**   | 4    | 显示最低值项（条件 1 中指定的项数）。   |
| **xlBottom10Percent** | 6    | 显示最低值项（条件 1 中指定的百分数）。 |
| **xlFilterCellColor** | 8    | 单元格颜色                              |
| **xlFilterDynamic**   | 11   | 动态筛选                                |
| **xlFilterFontColor** | 9    | 字体颜色                                |
| **xlFilterIcon**      | 10   | 筛选图标                                |
| **xlFilterValues**    | 7    | 筛选值                                  |
| **xlOr**              | 2    | 条件 1 和条件 2 的逻辑或。              |
| **xlTop10Items**      | 3    | 显示最高值项（条件 1 中指定的项数）。   |
| **xlTop10Percent**    | 5    | 显示最高值项（条件 1 中指定的百分数）。 |

注：使用xlAnd和xlOr可将参数Criteria1和Criteria2组合成复合筛选条件。

参数Criteria2是可选的，指定第2个筛选条件。与Criteria1和Operator组成复合筛选条件。

参数VisibleDropDown是可选的，如果其值为True(默认值)，则显示筛选字段自动筛选的下拉按钮。如果为False，则隐藏筛选字段自动筛选的下拉按钮。

提示：在excel中执行/取消筛选的快捷键为Ctrl+Shift+L

## 4.31 多条件筛选

通过多次使用AutoFilter方法，能够实现数据列表的多条件筛选功能。

如下示例代码通过两次筛选，检索"性别"字段内容为"男"并且"年龄"字段为22~26的记录。

```vb
Sub Filter_MoreCriteria()
    Application.ScreenUpdating = False
    With Worksheets("Sheet1")
        If .FilterMode = True Then .ShowAllData
        .Range("A1").AutoFilter Field:=2, Criteria1:="男"
        .Range("A1").AutoFilter Field:=3, Criteria1:=">=22", _
            Operator:=xlAnd, Criteria2:="<=26"
    End With
    Application.ScreenUpdating = True
End Sub
```

代码中使用FilterMode属性判断工作表是否处于筛选模式，如果是则显示当前筛选列表的所有数据。当工作表包含已筛选序列且该序列中含有隐藏行时，FilterMode属性的值为True。使用ShowAllData方法将显示所有数据，使当前筛选列表的所有行均可见。代码执行效果如下：

## 4.32获取符合筛选条件的记录数

在对工作表列表区域进行自动筛选时，在状态栏中将显示符合筛选条件的记录条数信息。

通常，自动筛选的结果由多个不连续的区域组成，这些单元格区域的行数即为记录数。

```vb
Sub GetFilterRecordCount()
    Dim rngRange As Range
    Dim i As Long
    Dim lngCount As Long
    Dim lngAllCount As Long
    With ActiveSheet
        If .FilterMode Then
            lngAllCount = .AutoFilter.Range.Rows.Count - 1'获取工作表筛选区域的记录总数
            Set rngRange = .AutoFilter.Range.SpecialCells(xlCellTypeVisible)'获取工作表筛选后筛选区域中的可视区域。
            For i = 1 To rngRange.Areas.Count
                lngCount = lngCount + rngRange.Areas(i).Rows.Count
            Next i
            Set rngRange = Nothing
            MsgBox "在 " & lngAllCount & " 条记录中找到 " & _
                lngCount - 1 & " 个"
        End If
    End With
End Sub
```

AutoFilter对象的Range属性返回工作表应用自动筛选的区域。

Areas属性返回一个Areas集合，代表多重选定区域中的所有区域。Areas集合包含选定区域内的每一个离散的连续单元格区域的Range对象。

## 4.33 判断筛选结果是否为空

```vb
Sub FilterIsEmpty()
    With ActiveSheet.AutoFilter.Range.SpecialCells _
        (xlCellTypeVisible)
        If .Areas.Count = 1 And .Rows.Count = 1 Then
            MsgBox "筛选结果为空"
        End If
    End With
End Sub
```

## 4.34复制自动筛选后的数据区域

```vb
Sub CopyFilterResult()
    With Worksheets("Sheet1")
        If .FilterMode Then
            .AutoFilter.Range.SpecialCells(xlCellTypeVisible).Copy _
                Worksheets("Sheet2").Range("A1")
        End If
    End With
End Sub
```

## 4.35 使用删除重复项获取不重复记录

利用单元格区域的RemoveDuplicates方法可以去除重复值，从而获取不重复值列表，如下图所示。

![image-20200219193619295](C:\Users\admin\AppData\Roaming\Typora\typora-user-images\image-20200219193619295.png)

示例代码如下。

```vb
Sub RemoveDuplicates()
    Range("A1").CurrentRegion.Copy Range("E1")
    Range("E1").CurrentRegion.RemoveDuplicates Columns:=1, Header:=xlYes'获取第1列不重复对应的所有记录，区域包括标题行
End Sub
```

RemoveDuplicates方法语法格式如下。

```
RemoveDuplicates(Columns, Header)
```

参数Columns是必需的，指定包含重复信息的列的索引数组。当有多列时，可以使用Array(1,2)的形式表示。

参数Header是可选的，指定第一行是否包含标题信息。默认值为xlNo；如果希望Excel自动判定，则应指定为XlGuess。

这里在介绍下Range对象的CurrentRegion属性。该属性返回包含指定单元格在内的一个连续的矩形区域，例如

```
Range("B5").CurrentRegion.Select
```

执行这行代码后的效果如下图所示。

而Worksheet对象的UsedRange属性返回工作表中已经使用的单元格围成的矩形区域，例如

```
Activesheet.UsedRange.Select '选中活动工作表中已经使用的单元格区域
```



## 4.41 数据排序

在如图1所示数据列表中，需要按总成绩从高到低进行排序，示例代码如下。

```vb
Sub SortDemo()
    Range("A1").Sort key1:="总成绩", order1:=xlDescending, _
        Header:=xlYes
End Sub
```

运行SortDemo过程，排序结果如图2所示。

Range对象的Sort方法对区域进行排序，其语法格式如下。

```
Sort(Key1, Order1, Key2, Type, Order2, Key3, Order3, Header, OrderCustom, MatchCase, Orientation, SortMethod, DataOption1, DataOption2, DataOption3)
```

其中，参数Key1、Key2和Key3是可选的，分别指定第1排序字段、第2排序字段、第3排序字段，作为区域名称(字符串)或Range对象，以确定要排序的值。

参数Order1、Order2和Order3是可选的，其值可为xlAscending，按升序对指定字段排序(默认值)；或者是xlDescending，按降序对指定字段排序。

参数Header是可选的，用于指定第1行是否包含标题信息，其值可为xlGuess，表示由Excel确定是否有标题；xlNo，表示不包含标题(默认值)；xlYes，表示包含标题。

Range对象的Sort方法最多可以指定3个排序字段，如下示例代码对图1所示数据集以“总成绩”、“学科1”和“学科2”分别为第1字段、第2字段和第3字段进行排序，排序结果如图3所示。

```vb
Sub SortDemoA()
    Range("A1").Sort key1:="总成绩", order1:=xlDescending, _
        key2:="学科1", order2:=xlDescending, _
        key3:="学科2", order3:=xlDescending, _
        Header:=xlYes
End Sub
```

## 4.42 多关键字排序

使用Range对象的Sort方法对区域进行排序时，同时最多只能指定3个关键字，当需要按照超过3个关键字对区域进行排序时，可以通过多次执行Sort方法实现。需要注意的是，在排序时应按照各关键字的倒叙顺序。例如，如果按照A→B→C→D的关键字顺序进行排序，则应按D→C→B→A的顺序执行Sort方法。

图1 带排序数据

如图1所示数据表中，需要按"总成绩"、"基础知识"、"教育学"和"心理学"的成绩降序排列，实例代码如下。

```vb
Sub SortByKeysA()
    With Range("A1")
        .Sort Key1:="心理学", order1:=xlDescending, Header:=xlYes
        .Sort Key1:="教育学", order1:=xlDescending, Header:=xlYes
        .Sort Key1:="基础知识", order1:=xlDescending, Header:=xlYes
        .Sort Key1:="总成绩", order1:=xlDescending, Header:=xlYes
    End With
End Sub
```

运行以上过程，结果如图2所示。



使用Range对象的Sort方法对区域进行超过3个关键字排序时，需要多次执行Sort方法，而通过Worksheet对象的Sort方法则可以一次完成。如下示例代码实现与上面示例代码相同的排序功能。

```vb
Sub MoreKeySort()
    With ActiveSheet.Sort.SortFields
        .Clear
        .Add Key:=Range("G1"), SortOn:=xlSortOnValues, Order:=xlDescending
        .Add Key:=Range("B1"), SortOn:=xlSortOnValues, Order:=xlDescending
        .Add Key:=Range("C1"), SortOn:=xlSortOnValues, Order:=xlDescending
        .Add Key:=Range("D1"), SortOn:=xlSortOnValues, Order:=xlDescending
    End With
    With ActiveSheet.Sort
        .SetRange Range("A1").CurrentRegion
        .Header = xlYes
        .Apply
    End With
End Sub
```

第3行代码清除工作表所有的SortFields对象。

第4~7行分别在Sort对象中添加SortFields对象。SortFields对象的Add方法创建新的排序字段，并返回SortFields对象，其语法格式如下。

`Add(Key, SortOn, Order, CustomOrder,  DataOption)`

该方法的各参数分别对应于Range对象Sort方法的参数。

第10行代码指定Sort对象的排序区域。

第11行代码指定排序区域包含标题。

第12行代码应用工作表排序。

## 4.43 自定义排序

在图1中所示的数据集中，如果希望按单元格区域E2:E6所列序列进行排序，需要先使用AddCustomList方法为应用程序添加自定义序列，示例代码如下。



```vb
Sub SortByLists()
    Dim avntList As Variant, lngNum As Long
    avntList = Range("E2:E6")
    Application.AddCustomList avntList
    lngNum = Application.GetCustomListNum(avntList)
    Range("A1").Sort Key1:=Range("A1"), _
        Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=lngNum + 1
    Application.DeleteCustomList lngNum
End Sub
```

第4行代码通过Application对象的AddCustomList方法为应用程序添加一个自定义序列。AddCustomList方法为自定义自动填充(或自定义排序)添加自定义列表，其语法格式如下。

`AddCustomList(ListArray, ByRow)`

其中，参数ListArray是必需的，可以为字符串数组或Range对象。参数ByRow是可选的，仅当 *ListArray* 为 **Range** 对象时使用。如果为 **True**，则使用区域中的每一行创建自定义列表；如果为  **False**，则使用区域中的每一列创建自定义列表。如果省略该参数，并且区域中的行数比列数多（或者行数与列数相等），则 Microsoft Excel  使用区域中的每一列创建自定义列表。如果省略该参数，并且区域中的列数比行数多，则 Microsoft Excel 使用区域中的每一行创建自定义列表。

如果要添加的列表已经存在，则AddCustomList方法不起作用。

第5行返回avnList数组在自定义序列中的序号。

第6行使用Sort方法对当前数据排序，其中Sort的参数指定了第1关键字Key1，默认为升序排列，同时设置包含标题，并且指定按新添加的自定义序列索引号排序。

注：参数OrderCustom指定在自定义排序次序列表中的基于1的整数偏移，在指定该参数时需在自定义序列号基础上加1。

第9行代码使用DeleteCustomList方法删除新添加的自定义序列。

## 利用Array函数向单元格填加内容

```vb
Range("A1:B1")=Array("序号","姓名")
```

Array函数返回一个包含数组的Variant，其语法格式如下。

`Array(arglist)`

arglist参数是一个用逗号隔开的元素值列表，这些值用于给数组的各元素赋值。



# 第6章 使用Shape对象

在Excel中Shape对象代表绘图层中的对象，如自选图形、任意多边形、OLE对象(Object Linking and Embedding，对象连接与嵌入)或图形。

![1](D:\微信公众号\6.1\1.JPG)

## 6.1 遍历工作表中的Shape对象

```vb
Sub ForNextAllShapes()
    Dim intRow As Integer
    Dim i As Integer
    Dim strShapeTypeConst As String
    intRow = 2
    With Sheets("Shape对象").Shapes
        For i = 1 To .Count
            With .Range(i)
                Sheets("统计").Cells(intRow, 1) = i
                Sheets("统计").Cells(intRow, 2) = .Name
                Sheets("统计").Cells(intRow, 3) = .Type
                Sheets("统计").Cells(intRow, 4) = .AutoShapeType
                If .Type = 1 Then'判断对象是否为自选图形
                    Select Case .AutoShapeType'判断自选图形对象的类型
                        Case 1
                            strShapeTypeConst = "矩形"
                        Case 5
                            strShapeTypeConst = "圆角矩形"
                        Case 92
                            strShapeTypeConst = "五角星"
                    End Select
                Else
                    Select Case .Type
                        Case 5
                            strShapeTypeConst = "任意多边形"
                        Case 8
                            strShapeTypeConst = "窗体控件"
                        Case 9
                            strShapeTypeConst = "线条"
                        Case 12
                            strShapeTypeConst = "OLE 控件对象"
                        Case 13
                            strShapeTypeConst = "图片"
                    End Select
                End If
            End With
            Sheets("统计").Cells(intRow, 5) = strShapeTypeConst
            intRow = intRow + 1
        Next i
    End With
End Sub
```

第7行代码使用了Shapes对象的Count属性返回"Shapes对象"工作表中Shape对象的总数量。

第8行代码使用Shapes对象的Range属性返回Shape对象集合中图形的一个子集，其语法格式如下。

`expression.Range(Index)`

其中，参数Index可以是指定图形索引号的整数，或者是图形名称的字符串，也可以是包含整数或字符串的数组。

注：此处的Range是Shape对象的属性，返回一个ShapeRange对象，它代表Shapes集合中形状的子集，不同于工作表中的Range对象。

第9~12行代码将Shape对象的序号、名称、Type属性值和AutoShapeType属性值写入"统计"工作表中。

第13行代码使用Shape对象的Type属性判断对象是否为自选图形。

Shape对象的Type属性返回或设置一个MsoShapeType值，该值代表Shape对象的类型，MsoShapeType常量值与说明如下表所示。

| 名称                     | 值   | 说明              |
| ------------------------ | ---- | ----------------- |
| **msoAutoShape**         | 1    | 自选图形。        |
| **msoCallout**           | 2    | 标注。            |
| **msoCanvas**            | 20   | 画布。            |
| **msoChart**             | 3    | 图。              |
| **msoComment**           | 4    | 批注。            |
| **msoDiagram**           | 21   | 图表。            |
| **msoEmbeddedOLEObject** | 7    | 嵌入的 OLE 对象。 |
| **msoFormControl**       | 8    | 窗体控件。        |
| **msoFreeform**          | 5    | 任意多边形。      |
| **msoGroup**             | 6    | 组合。            |
| **msoIgxGraphic**        | 24   | SmartArt 图形     |
| **msoInk**               | 22   | 墨迹。            |
| **msoInkComment**        | 23   | 墨迹批注。        |
| **msoLine**              | 9    | 线条。            |
| **msoLinkedOLEObject**   | 10   | 链接 OLE 对象。   |
| **msoLinkedPicture**     | 11   | 链接图片。        |
| **msoMedia**             | 16   | 媒体。            |
| **msoOLEControlObject**  | 12   | OLE 控件对象。    |
| **msoPicture**           | 13   | 图片。            |
| **msoPlaceholder**       | 14   | 占位符。          |
| **msoScriptAnchor**      | 18   | 脚本定位标记。    |
| **msoShapeTypeMixed**    | -2   | 混和形状类型。    |
| **msoTable**             | 19   | 表。              |
| **msoTextBox**           | 17   | 文本框。          |
| **msoTextEffect**        | 15   | 文本效果。        |

Shape对象的AutoShapeType属性返回或设置一个MsoAutoShapeType值，该值指定Shape或ShapeRange对象的类型，该对象必须是自选图形，不能是直线、任意多边形图形或连接符。MsoAutoShapeType常量值与说明如下表所示。

| 名称                                         | 值   | 说明                                                       |
| -------------------------------------------- | ---- | ---------------------------------------------------------- |
| **msoShape16pointStar**                      | 94   | 十六角星。                                                 |
| **msoShape24pointStar**                      | 95   | 二十四角星。                                               |
| **msoShape32pointStar**                      | 96   | 三十二角星。                                               |
| **msoShape4pointStar**                       | 91   | 四角星。                                                   |
| **msoShape5pointStar**                       | 92   | 五角星。                                                   |
| **msoShape8pointStar**                       | 93   | 八角星。                                                   |
| **msoShapeActionButtonBackorPrevious**       | 129  | **“后退”**或**“上一个”**按钮。支持鼠标单击和鼠标移过操作。 |
| **msoShapeActionButtonBeginning**            | 131  | **“开始”**按钮。支持鼠标单击和鼠标移过操作。               |
| **msoShapeActionButtonCustom**               | 125  | 不带默认图片或文本的按钮。支持鼠标单击和鼠标移过操作。     |
| **msoShapeActionButtonDocument**             | 134  | **“文档”**按钮。支持鼠标单击和鼠标移过操作。               |
| **msoShapeActionButtonEnd**                  | 132  | **“结束”**按钮。支持鼠标单击和鼠标移过操作。               |
| **msoShapeActionButtonForwardorNext**        | 130  | **“前进”**或**“下一个”**按钮。支持鼠标单击和鼠标移过操作。 |
| **msoShapeActionButtonHelp**                 | 127  | **帮助**按钮。支持鼠标单击和鼠标移过操作。                 |
| **msoShapeActionButtonHome**                 | 126  | **“主页”**按钮。支持鼠标单击和鼠标移过操作。               |
| **msoShapeActionButtonInformation**          | 128  | **“信息”**按钮。支持鼠标单击和鼠标移过操作。               |
| **msoShapeActionButtonMovie**                | 136  | **“影片”**按钮。支持鼠标单击和鼠标移过操作。               |
| **msoShapeActionButtonReturn**               | 133  | **“返回”**按钮。支持鼠标单击和鼠标移过操作。               |
| **msoShapeActionButtonSound**                | 135  | **“声音”**按钮。支持鼠标单击和鼠标移过操作。               |
| **msoShapeArc**                              | 25   | 弧形。                                                     |
| **msoShapeBalloon**                          | 137  | 气球。                                                     |
| **msoShapeBentArrow**                        | 41   | 带 90 度圆角的箭头。                                       |
| **msoShapeBentUpArrow**                      | 44   | 带 90 度直角的箭头。默认情况下上指。                       |
| **msoShapeBevel**                            | 15   | 凹凸效果。                                                 |
| **msoShapeBlockArc**                         | 20   | 空心弧。                                                   |
| **msoShapeCan**                              | 13   | 圆柱形。                                                   |
| **msoShapeChevron**                          | 52   | V 形。                                                     |
| **msoShapeCircularArrow**                    | 60   | 带 180 度圆角的箭头。                                      |
| **msoShapeCloudCallout**                     | 108  | 云形标注。                                                 |
| **msoShapeCross**                            | 11   | 十字形。                                                   |
| **msoShapeCube**                             | 14   | 立方。                                                     |
| **msoShapeCurvedDownArrow**                  | 48   | 上弧形箭头。                                               |
| **msoShapeCurvedDownRibbon**                 | 100  | 下凸弯带形横幅。                                           |
| **msoShapeCurvedLeftArrow**                  | 46   | 右弧形箭头。                                               |
| **msoShapeCurvedRightArrow**                 | 45   | 左弧形箭头。                                               |
| **msoShapeCurvedUpArrow**                    | 47   | 下弧形箭头。                                               |
| **msoShapeCurvedUpRibbon**                   | 99   | 上凸弯带形。                                               |
| **msoShapeDiamond**                          | 4    | 菱形。                                                     |
| **msoShapeDonut**                            | 18   | 环形。                                                     |
| **msoShapeDoubleBrace**                      | 27   | 双大括号。                                                 |
| **msoShapeDoubleBracket**                    | 26   | 双括号。                                                   |
| **msoShapeDoubleWave**                       | 104  | 双波形。                                                   |
| **msoShapeDownArrow**                        | 36   | 下箭头。                                                   |
| **msoShapeDownArrowCallout**                 | 56   | 带下箭头的标注。                                           |
| **msoShapeDownRibbon**                       | 98   | 中心区域位于弯带末端下方的弯带形。                         |
| **msoShapeExplosion1**                       | 89   | 爆炸形。                                                   |
| **msoShapeExplosion2**                       | 90   | 爆炸形。                                                   |
| **msoShapeFlowchartAlternateProcess**        | 62   | 其他过程流程图符号。                                       |
| **msoShapeFlowchartCard**                    | 75   | 资料卡流程图符号。                                         |
| **msoShapeFlowchartCollate**                 | 79   | 对照流程图符号。                                           |
| **msoShapeFlowchartConnector**               | 73   | 联系流程图符号。                                           |
| **msoShapeFlowchartData**                    | 64   | 数据流程图符号。                                           |
| **msoShapeFlowchartDecision**                | 63   | 决策流程图符号。                                           |
| **msoShapeFlowchartDelay**                   | 84   | 延期流程图符号。                                           |
| **msoShapeFlowchartDirectAccessStorage**     | 87   | 磁鼓流程图符号。                                           |
| **msoShapeFlowchartDisplay**                 | 88   | 显示流程图符号。                                           |
| **msoShapeFlowchartDocument**                | 67   | 文档流程图符号。                                           |
| **msoShapeFlowchartExtract**                 | 81   | 摘录流程图符号。                                           |
| **msoShapeFlowchartInternalStorage**         | 66   | 内部贮存流程图符号。                                       |
| **msoShapeFlowchartMagneticDisk**            | 86   | 磁盘流程图符号。                                           |
| **msoShapeFlowchartManualInput**             | 71   | 手动输入流程图符号。                                       |
| **msoShapeFlowchartManualOperation**         | 72   | 手动操作流程图符号。                                       |
| **msoShapeFlowchartMerge**                   | 82   | 合并流程图符号。                                           |
| **msoShapeFlowchartMultidocument**           | 68   | 多文档流程图符号。                                         |
| **msoShapeFlowchartOffpageConnector**        | 74   | 离页连接符流程图符号。                                     |
| **msoShapeFlowchartOr**                      | 78   | “或者”流程图符号。                                         |
| **msoShapeFlowchartPredefinedProcess**       | 65   | 预定义过程流程图符号。                                     |
| **msoShapeFlowchartPreparation**             | 70   | 准备流程图符号。                                           |
| **msoShapeFlowchartProcess**                 | 61   | 过程流程图符号。                                           |
| **msoShapeFlowchartPunchedTape**             | 76   | 资料带流程图符号。                                         |
| **msoShapeFlowchartSequentialAccessStorage** | 85   | 磁带流程图符号。                                           |
| **msoShapeFlowchartSort**                    | 80   | 排序流程图符号。                                           |
| **msoShapeFlowchartStoredData**              | 83   | 库存数据流程图符号。                                       |
| **msoShapeFlowchartSummingJunction**         | 77   | 汇总连接流程图符号。                                       |
| **msoShapeFlowchartTerminator**              | 69   | 终止流程图符号。                                           |
| **msoShapeFoldedCorner**                     | 16   | 折角形。                                                   |
| **msoShapeHeart**                            | 21   | 心形。                                                     |
| **msoShapeHexagon**                          | 10   | 六边形。                                                   |
| **msoShapeHorizontalScroll**                 | 102  | 横卷形。                                                   |
| **msoShapeIsoscelesTriangle**                | 7    | 等腰三角形。                                               |
| **msoShapeLeftArrow**                        | 34   | 左箭头。                                                   |
| **msoShapeLeftArrowCallout**                 | 54   | 带左箭头的标注。                                           |
| **msoShapeLeftBrace**                        | 31   | 左大括号。                                                 |
| **msoShapeLeftBracket**                      | 29   | 左括号。                                                   |
| **msoShapeLeftRightArrow**                   | 37   | 左右双向箭头。                                             |
| **msoShapeLeftRightArrowCallout**            | 57   | 带左右双向箭头的标注。                                     |
| **msoShapeLeftRightUpArrow**                 | 40   | 左右上三向箭头。                                           |
| **msoShapeLeftUpArrow**                      | 43   | 左上双向箭头。                                             |
| **msoShapeLightningBolt**                    | 22   | 闪电形。                                                   |
| **msoShapeLineCallout1**                     | 109  | 带边框和水平标注线的标注。                                 |
| **msoShapeLineCallout1AccentBar**            | 113  | 带水平强调线的标注。                                       |
| **msoShapeLineCallout1BorderandAccentBar**   | 121  | 带边框和水平强调线的标注。                                 |
| **msoShapeLineCallout1NoBorder**             | 117  | 带水平线的标注。                                           |
| **msoShapeLineCallout2**                     | 110  | 带对角直线的标注。                                         |
| **msoShapeLineCallout2AccentBar**            | 114  | 带对角标注线和强调线的标注。                               |
| **msoShapeLineCallout2BorderandAccentBar**   | 122  | 带边框、对角直线和强调线的标注。                           |
| **msoShapeLineCallout2NoBorder**             | 118  | 不带边框和对角标注线的标注。                               |
| **msoShapeLineCallout3**                     | 111  | 带倾斜线的标注。                                           |
| **msoShapeLineCallout3AccentBar**            | 115  | 带倾斜标注线和强调线的标注。                               |
| **msoShapeLineCallout3BorderandAccentBar**   | 123  | 带边框、倾斜标注线和强调线的标注。                         |
| **msoShapeLineCallout3NoBorder**             | 119  | 不带边框和倾斜标注线的标注。                               |
| **msoShapeLineCallout4**                     | 112  | 带 U 型标注线段的标注。                                    |
| **msoShapeLineCallout4AccentBar**            | 116  | 带强调线和 U 型标注线段的标注。                            |
| **msoShapeLineCallout4BorderandAccentBar**   | 124  | 带边框、强调线和 U 型标注线段的标注。                      |
| **msoShapeLineCallout4NoBorder**             | 120  | 不带边框和 U 型标注线段的标注。                            |
| **msoShapeMixed**                            | -2   | 只返回值，表示其他状态的组合。                             |
| **msoShapeMoon**                             | 24   | 新月形。                                                   |
| **msoShapeNoSymbol**                         | 19   | 禁止符。                                                   |
| **msoShapeNotchedRightArrow**                | 50   | 燕尾形右箭头。                                             |
| **msoShapeNotPrimitive**                     | 138  | 不支持。                                                   |
| **msoShapeOctagon**                          | 6    | 八边形。                                                   |
| **msoShapeOval**                             | 9    | 椭圆形。                                                   |
| **msoShapeOvalCallout**                      | 107  | 椭圆形标注。                                               |
| **msoShapeParallelogram**                    | 2    | 平行四边形。                                               |
| **msoShapePentagon**                         | 51   | 五边形。                                                   |
| **msoShapePlaque**                           | 28   | 缺角矩形。                                                 |
| **msoShapeQuadArrow**                        | 39   | 四向箭头。                                                 |
| **msoShapeQuadArrowCallout**                 | 59   | 带四向箭头的标注。                                         |
| **msoShapeRectangle**                        | 1    | 矩形。                                                     |
| **msoShapeRectangularCallout**               | 105  | 矩形标注。                                                 |
| **msoShapeRegularPentagon**                  | 12   | 五边形。                                                   |
| **msoShapeRightArrow**                       | 33   | 右箭头。                                                   |
| **msoShapeRightArrowCallout**                | 53   | 带右箭头的标注。                                           |
| **msoShapeRightBrace**                       | 32   | 右大括号。                                                 |
| **msoShapeRightBracket**                     | 30   | 右括号。                                                   |
| **msoShapeRightTriangle**                    | 8    | 直角三角形。                                               |
| **msoShapeRoundedRectangle**                 | 5    | 圆角矩形。                                                 |
| **msoShapeRoundedRectangularCallout**        | 106  | 圆角矩形标注。                                             |
| **msoShapeSmileyFace**                       | 17   | 笑脸。                                                     |
| **msoShapeStripedRightArrow**                | 49   | 尾部带条纹的右箭头。                                       |
| **msoShapeSun**                              | 23   | 太阳。                                                     |
| **msoShapeTrapezoid**                        | 3    | 梯形。                                                     |
| **msoShapeUpArrow**                          | 35   | 上箭头。                                                   |
| **msoShapeUpArrowCallout**                   | 55   | 带上箭头的标注。                                           |
| **msoShapeUpDownArrow**                      | 38   | 上下双向箭头。                                             |
| **msoShapeUpDownArrowCallout**               | 58   | 带上下双向箭头的标注。                                     |
| **msoShapeUpRibbon**                         | 97   | 中心区域位于弯带末端上方的弯带形横幅。                     |
| **msoShapeUTurnArrow**                       | 42   | U 型箭头。                                                 |
| **msoShapeVerticalScroll**                   | 101  | 竖卷形。                                                   |
| **msoShapeWave**                             | 103  | 波形。                                                     |

还可以使用For Each循环结构遍历Shapes对象集合中的Shape对象。

```vb
Sub ForEachAllShapes()
    Dim objShp As Shape
    For Each objShp In Sheets("Shape对象").Shapes
        Debug.Print objShp.Name, Tab(30), objShp.Type, _
            Tab(60), objShp.AutoShapeType
    Next objShp
    Set objShp = Nothing
End Sub
```

Debug.Print会在立即窗口中输出

## 6.2在工作表中快速添加Shape对象

```vb
Sub InsertShape()
    Dim intxOffset As Integer
    Dim intyOffset As Integer
    Dim intRow As Integer
    Dim objLine As LineFormat'声明线条对象
    Dim objFreeForm As FreeformBuilder'声明为任意多边形对象
    intyOffset = 50
    intxOffset = 50
    With Sheets("数据")
        For intRow = 2 To 11
            Set objLine = Sheets("绘图区").Shapes.AddLine( _
                .Cells(intRow, 1) + intxOffset, _
                .Cells(intRow, 2) + intyOffset, _
                .Cells(intRow + 1, 1) + intxOffset, _
                .Cells(intRow + 1, 2) + intyOffset).Line
        objLine.Weight = .Cells(intRow, 3)'设置线条粗线
            objLine.ForeColor.RGB = .Cells(intRow, 4)'设置线条前景填充色
        Next intRow
    End With
    intxOffset = intxOffset + 300
    intRow = 2
    With Sheets("数据")
        Set objFreeForm = Sheets("绘图区").Shapes.BuildFreeform( _
            msoEditingAuto, _
            .Cells(intRow, 1) + intxOffset, _
            .Cells(intRow, 2) + intyOffset)
        For intRow = 3 To 12
            objFreeForm.AddNodes msoSegmentLine, msoEditingAuto, _
                .Cells(intRow, 1) + intxOffset, _
                .Cells(intRow, 2) + intyOffset
        Next intRow
        objFreeForm.ConvertToShape
    End With
    intxOffset = intxOffset + 300
    Sheets("绘图区").Shapes.AddShape(msoShape5pointStar, _
        intxOffset, intyOffset, 266.3, 253.26).Select
    Selection.ShapeRange.Fill.ForeColor.RGB = RGB(255, 0, 0)
    Set objLine = Nothing
    Set objFreeForm = Nothing
End Sub
```

结果：

![image-20200223224532079](C:\Users\admin\AppData\Roaming\Typora\typora-user-images\image-20200223224532079.png)



第2行和第3行代码声明Integer类型变量用于保存图形相对于**文档左上角(即A1单元格左上角)**的偏移量。

第5行声明变量为LineFormat对象，代表线条和箭头格式。

第11行将AddLine方法返回的线条对象赋值给变量objLine，Shapes对象的AddLine方法创建并返回一个线条类型的对象，AddLine方法的语法如下。

`AddLine(BeginX, BeginY, EndX, EndY)`

这四个参数都是必选参数。其中BeginX和BeginY为线条起点相对于文档左上角的水平和垂直坐标；EndX和EndY为线条终点相对于文档左上角的水平和垂直坐标。

在数据表中，"X坐标"是水平坐标，"Y坐标"是垂直坐标。

第35行代码用来创建右侧的五角星对象。Shapes对象的AddShape方法返回一个Shape对象，该对象代表新添加的自选图形，其语法格式如下。

`AddShape(Type,Left,Top,Width,Height)`

AddShape方法的5个参数都是必需的，参数Type用来指定要添加的自选图形的类型，后面4个参数都以磅为单位。

ShapeRange对象代表形状区域，Fill属性返回指定形状的FillFormat对象或指定图表的ChartFillFormat对象，这两种对象包含形状或图表的填充格式属性，可以设置Shape对象的填充色。

## 6.3 组合多个Shape对象

如果在工作表中存在多个Shape对象时，在工作表插入或删除单元格、改变行高或列宽时，可能改变形状的大小及它们之间的相对位置。设置自选图形的组合，可以保持组合内自选图形之间的相对位置不发生变化，并且对组合后图形的操作如同处理单个Shape对象，设置对象附加到单元格的方式为自由浮动，可以保持对象大小和位置固定不变，示例代码如下。

```vb
Sub GroupShapes()
    Dim i As Integer
    Dim astrLineName(1 To 10) As String'声明数组用于保存10条线段的名称。
    For i = 1 To 10
        astrLineName(i) = "直接连接符 " & i
    Next i
    With Sheet1.Shapes
        .Range(astrLineName()).Group.Placement = xlFreeFloating
        .Range(Array("任意多边形 11", _
            "五角星 12")).Group.Placement = xlFreeFloating
    End With
    Sheet1.Shapes.SelectAll
End Sub
```

第8行使用ShapeRange对象的Group方法将10条线条组合到一起(其中astrLineName数组作为Range属性的参数)，并且设置组合后的Shape对象不会随单元格移动或调整大小。

Group方法将指定区域中的形状组合在一起，并返回一个代表组合形状的Shape对象。

Placement属性返回或设置一个XlPlacement值，代表对象附加到单元格单元格的方式，XlPlacement常量值与说明如下表所示。

| 名称               | 值   | 说明                         |
| ------------------ | ---- | ---------------------------- |
| **xlFreeFloating** | 3    | 对象自由浮动。               |
| **xlMove**         | 2    | 对象随单元格移动。           |
| **xlMoveAndSize**  | 1    | 对象随单元格移动和调整大小。 |

注：在Excel中将组合形状视为单个Shape对象，组合或取消组合形状时，Shapes对象集合中的对象个数将发生变化，索引号也会变。因此Range属性的参数应尽量采用Shape对象的名称，避免使用其索引号。



![image-20200226194714643](C:\Users\admin\AppData\Roaming\Typora\typora-user-images\image-20200226194714643.png)

![image-20200226194952878](C:\Users\admin\AppData\Roaming\Typora\typora-user-images\image-20200226194952878.png)

## 批量添加PDF文件

如图1所示，"BOM-01.xlsx"工作簿中的Sheet1工作表根据B列图号单元格中的内容，在图2所示的文件夹中找到对应的PDF文件，然后嵌入到相应的N列，双击N列中所示的图标，会打开PDF文件，是源文件的副本，即删除源文件，也可以打开N列的文件。

![image-20200226205333296](C:\Users\admin\AppData\Roaming\Typora\typora-user-images\image-20200226205333296.png)

图1 根据图号单在N列嵌入对应的PDF文件

![image-20200226210024505](C:\Users\admin\AppData\Roaming\Typora\typora-user-images\image-20200226210024505.png)

图2 PDF文件

```vb
Sub 导入文件()
    Application.ScreenUpdating = False'禁止屏幕更新
    Application.DisplayAlerts = False'禁止弹出对话框
    
    Dim fil As String, fn As String
    Dim wb As Workbook
    Dim sht As Worksheet
    Dim RWidth As Long, RHeight As Long
    Dim Obj As Object
    
    RWidth = 40
    RHeight = 60
    Set wb = Workbooks.Open(ThisWorkbook.Path & "\BOM-01.xlsx")
    Set sht = wb.Worksheets(1)
    sht.Columns("N:N").ColumnWidth = RWidth
    Dim Str1 As String
    Dim FileName As String
    FileName = Dir(ThisWorkbook.Path & "\PDF文件\*.pdf")
    Dim i As Long, IRow As Long
    IRow = sht.Range("B10000").End(xlUp).Row
    
    Do While FileName <> ""
        For i = 4 To IRow
            Str1 = Trim(sht.Cells(i, 2).Value)
            If InStr(FileName, Str1) And Str1 <> "" Then
                sht.Cells(i, "N").RowHeight = RHeight
                sht.Cells(i, "N").Select
                fn = ThisWorkbook.Path & "\PDF文件\" & FileName
                sht.OLEObjects.Add FileName:=fn, _
                    link:=False, _
                    DisplayAsIcon:=True, _
                    IconFileName:="C:\windows\Installer\{AC76BA86-1033-FFFF-7760-0E0F06755100}\_PDFFile.ico", _
                    iconindex:=0, _
                    iconlabel:=fn
                Exit For
            End If
        Next
        
        FileName = Dir '用dir函数取得其他文件名，并赋给变量
    Loop
    wb.Save
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
```

OLEObjects.Add方法向工作表中添加新的 OLE 对象。其语法格式如下。

`表达式.Add(ClassType, FileName,  Link, DisplayAsIcon, IconFileName, IconIndex,  IconLabel, Left, Top, Width, Height)`

各参数说明如下表所示。

| **名称**        | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *ClassType*     | 可选          | **Variant**  | （必须指定 *ClassType* 或  *FileName*）。一个字符串，包含要创建的对象的程序标识符。如果指定了 *ClassType* 参数，则忽略  *FileName* 和 *Link*。 |
| *FileName*      | 可选          | **Variant**  | （必须指定 *ClassType* 或 *FileName*）。一个字符串，指定用于创建 OLE  对象的文件。 |
| *Link*          | 可选          | **Variant**  | 如果为 **True**，则让基于 *FileName* 的新 OLE  对象链接到该文件。如果该对象未链接到文件，则该对象被创建为文件副本。默认值是 **False**。 |
| *DisplayAsIcon* | 可选          | **Variant**  | 如果为 **True**，则以图标或正常图片方式显示新的 OLE 对象。如果该参数设置为  **True**，则可以使用 *IconFileName* 和 *IconIndex* 来指定图标。 |
| *IconFileName*  | 可选          | **Variant**  | 一个字符串，指定要显示的图标所在的文件。仅当 *DisplayAsIcon* 为 **True**  时，才使用该参数。如果不指定该参数，或文件中不包含图标，则使用 OLE 类的默认图标。 |
| *IconIndex*     | 可选          | **Variant**  | 图标文件中包含的图标数目。仅当 *DisplayAsIcon* 参数为 **True** 并且  *IconFileName* 参数引用包含图标的有效文件时，才使用该参数。如果由 *IconFileName*  参数指定的文件中不存在具有指定索引号的图标，则使用该文件中的第一个图标。 |
| *IconLabel*     | 可选          | **Variant**  | 一个字符串，指定在图标下方显示一个标签。仅当 *DisplayAsIcon* 为 **True**  时，才使用该参数。如果省略该参数，或者该参数为空字符串 ("")，则不显示任何标题。 |
| *Left*          | 可选          | **Variant**  | 以磅为单位给出新对象的初始坐标，该坐标是相对于工作表上单元格 A1  的左上角或图表的左上角的坐标。 |
| *Width*         | 可选          | **Variant**  | 以磅为单位给出新对象的初始大小。                             |