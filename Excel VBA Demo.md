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

# 2.裁剪图片

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

# 生成票号

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



# 四、Range操作

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

命名区域HighLightArea(示例文件已指定B2:H15单元格区域)

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

| 名称                      | 值   | 说明                         |
| :------------------------ | :--- | :--------------------------- |
| **xlValidateCustom**      | 7    | 使用任意公式验证数据有效性。 |
| **xlValidateDate**        | 4    | 日期值。                     |
| **xlValidateDecimal**     | 2    | 数值。                       |
| **xlValidateInputOnly**   | 0    | 仅在用户更改值时进行验证。   |
| **xlValidateList**        | 3    | 值必须存在于指定列表中。     |
| **xlValidateTextLength**  | 6    | 文本长度。                   |
| **xlValidateTime**        | 5    | 时间值。                     |
| **xlValidateWholeNumber** | 1    | 全部数值。                   |

参数Formula2指定数据验证公式的第二部分。仅当Operator为xlBetween或xlNotBetween时有效。

## 4.14 判断单元格公式是否存在错误

Excel公式返回的结果可能是一个错误的文本，包含#NULL、#DIV/0!、#VALUE！、#REF！、#NAME?、#NUM！和#N/A等。

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

