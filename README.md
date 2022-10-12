# fastexcel

![xxx](https://img.shields.io/badge/version-1.0.0-green) ![xxx](https://img.shields.io/badge/jdk-1.8-green)
![xxx](https://img.shields.io/badge/poi-5.0.0-green) ![xxx](https://img.shields.io/badge/spring-5.3.23-green)

## 1. 项目简介

fastexcel是基于 ([Apache POI](https://poi.apache.org/))工具组件。 它指在帮助我们通过简化、优雅、易懂的代码实现对Excel文件的快速读写。

## 2. 快速上手

### 2.1 快速写

```java
public class ExcelWriteTest {
    /**
      *  这里还支持嵌套对象、原生格式化、列宽、表头样式、表头批注、数据列验证等等
      *  更多写的功能待文档完善，或加群咨询
      *  https://blog.csdn.net/yyf314922957/article/details/111034523
     */
	@Test
	public void write() {
		Workbook workbook = new WorkbookBuilder(new SXSSFWorkbook())
			.setDefaultRowHeight(20)
			// 设置全局样式
			.matchingAll()
			.setFontHeight(12)
			.setFontName("微软雅黑")
			.setFillBackgroundColor(IndexedColors.WHITE1)
			.setVerticalAlignment(VerticalAlignment.CENTER)
			.setAlignment(HorizontalAlignment.CENTER)
			.setWrapText(true)
			.addCellStyle()
			// 匹配指定单元格设置样式
			.matchingCell(cell -> {
				// 只匹配分数列
				if (cell == null || cell.getColumnIndex() < 7) {
					return false;
				}
				// 小于60分的标红
				String cellValue = CellUtils.getCellValue(cell);
				if (NumberUtils.isCreatable(cellValue)) {
					return Double.parseDouble(cellValue) < 60D;
				}
				return false;
			})
			.setStrikeout(true)
			.setFontColor(IndexedColors.RED)
			.addCellStyle()
			.matching(cell -> {
				// 只匹配分数列
				if (cell == null || cell.getColumnIndex() < 7) {
					return false;
				}
				// 等于100分的标蓝
				String cellValue = CellUtils.getCellValue(cell);
				if (NumberUtils.isCreatable(cellValue)) {
					return Double.parseDouble(cellValue) == 100D;
				}
				return false;
			})
			.setFontColor(IndexedColors.BLUE)
			.addCellStyle()
			.createSheet("Sheet1")
			// 创建普通行。并且指定区域合并
			.createRow("班级概况")
			.addCellRange(0, 0, 0, 12)
			.merge()
			// 指定行映射对象
			.rowType(RowModel.class)
			// 创建表头，可以指定导出部分字段，不传代表导出所有
			.createHeader()
			.createRows(rowModels)
			.end()
			.build();
	}
}
```

### 2.2 快速读

```java
public class ExcelReadTest {
	
    /**
     * 这里只描述了如何读数据，很多时候我们读取数据发生错误时，
     * 如数据格式错误，需要将错误信息返回给用户
     * 这里提供了一种非常友好，使用也很简单的方式。
     * 就是将错误信息以批注的方式写回导入的excel，同时也支持根据业务传入自定义的错误
     */
	@Test
	public void read() {
		// 创建行映射器
		AnnotationBasedRowSetMapper<RowModel> rowSetMapper
                    = new AnnotationBasedRowSetMapper<>(RowModel.class);
		// 读取excel文件
		RowSetReader rowSetReader = RowSetReader.open(file.getInputStream());
		// 获取excel行数据
		for (RowSet rowSet : rowSetReader) {
			// 直接获取原始行数据
			String[] cellValues = rowSet.getRow().getCellValues();
			// 将原始行转换成对象
			MappingResult<RowModel> result = rowSetMapper.getMappingResult(rowSet);
			// 判断行转对象是否失败，如果没有错误。获取对象。
			if(!result.hasErrors()){
				ImportDTO target = result.getTarget();
			}
		}
	}
}
```

## 3. 最新版本

```xml

<dependency>
    <groupId>io.github.yyfcode</groupId>
    <artifactId>fastexcel</artifactId>
    <version>1.0.0</version>
</dependency>
```

## 4. 联系方式

有疑问请加钉钉群:2880006273，如果有人能协助完善文档、测试、修复bug以及后续迭代请加群联系。
