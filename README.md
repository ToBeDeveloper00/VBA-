# My Awesome Book

This file file serves as your book's preface, a great place to describe your book's content and ideas.

# 1.使用 QRMaker.ocx控件生成二维码

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

```VBA
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


​
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

```
Application.AskToUpdateLinks=False'关闭更新数据源链接的提示


Application.AskToUpdateLinks=True'将属性值恢复为默认状态
```

# 四、Range操作

## 4.2取得最后一个非空单元格

xlDown/xlToRight/xlToLeft/xlUp

```
Dim ERow as Long


Erow=Range("A" 
&
 Rows.Count).End(xlUp).Row
```

## 4.3 复制单元格区域

注意：使用PasteSpecial方法时指定xlPasteAll（粘贴全部），并不包括粘贴列宽

```
Sub CopyWithSameColumnWidths()


  Sheets("Sheet1").Range("A1").CurrentRegion.Copy


  With Sheets("Sheet2").Range("A1")


  .PasteSpecial xlPasteColumnWidths


  .PasteSpecial xlPasteAll


  End With


  Application.CutCopyMode = False


End Sub


​


Sheets("Sheet2").Range("A1").PasteSpecial Paste:=xlPasteValues '粘贴数值
```

## 4.9 设置字符格式

```
Sub CellCharacter()


  With Range("A1")


  .Clear


  .Value = "Y=X2+1"


  .Characters(4, 1).Font.Superscript = True'将第4个字符设置为上标


  .Characters(1, 1).Font.ColorIndex = 3


  .Font.Size = 20


  End With


End Sub
```

通过Range对象的Characters属性来操作指定的字符。Characters属性返回一个Characters对象，代表对象文字的字符区域。Characters属性的语法格式如下`Characters(Start, Length)`

## 4.10 单元格区域添加边框

使用Range对象的Borders集合可以快速地对单元格区域全部边框应用相同的格式。Range对象的BorderAround方法可以快速地为单元格区域添加外边框。

```
Sub AddBorders()


  Dim rngCell As Range


  Set rngCell = Range("B2:F8")


  With rngCell.Borders


  .LineStyle = xlContinuous'边框线条的样式


  .Weight = xlThin'设置边框线条粗细


  .ColorIndex = 5'设置边框线条颜色


  End With


  rngCell.BorderAround xlContinuous, xlMedium, 5'添加一个加粗外边框


  Set rngCell = Nothing


End Sub
```

在单元格区域中应用多种边框格式

```
Sub BordersIndexDemo()


  Dim rngCell As Range


  Set rngCell = Range("B2:F8")


  With rngCell.Borders(xlInsideHorizontal)'内部水平


  .LineStyle = xlDot


  .Weight = xlThin


  .ColorIndex = 5


  End With


  With rngCell.Borders(xlInsideVertical)'内部垂直


  .LineStyle = xlContinuous


  .Weight = xlThin


  .ColorIndex = 5


  End With


  rngCell.BorderAround xlContinuous, xlMedium, 5


  Set rngCell = Nothing


End Sub
```



