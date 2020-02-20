



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

关注微信公众号：**VBA168**
每天更新Excel VBA经典代码，祝你工作和学习更轻松！

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



