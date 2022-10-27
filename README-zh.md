# fastexcel

![xxx](https://img.shields.io/badge/version-1.0.0-green) ![xxx](https://img.shields.io/badge/jdk-1.8-green) ![xxx](https://img.shields.io/badge/poi-5.0.0-green) ![xxx](https://img.shields.io/badge/spring-5.3.23-green)

一个用于快速、方便地读写大的excel文件的Java库。在这个文件中您可以找到一些简单的示例，
您还可以查看一个使用这个库的Spring Boot应用程序：
[https://github.com/yyfcode/fastexcel-demo](https://github.com/yyfcode/fastexcel-demo) 。

## 案例

```java
public class ExcelWriteTest {

	@Test
	public void write() {
		Workbook workbook = WorkbookBuilder.builder()
			.setDefaultRowHeight(30)
			.setDefaultColumnWidth(30)
			.createSheet("Sheet 1")
			.matchingCell(cell -> {
				if (cell == null || cell.getColumnIndex() != 7) {
					return false;
				}
				String cellValue = CellUtils.getCellValue(cell);
				if (NumberUtils.isCreatable(cellValue)) {
					return Integer.parseInt(cellValue) < 60;
				}
				return false;
			})
			.setStrikeout(true)
			.setFontHeight(30)
			.setFontColor(IndexedColors.RED.getIndex())
			.addCellStyle()
			.rowType(Owner.class)
			.createHeader()
			.createRows(Arrays.asList(george, joe))
			.end()
			.build();
	}
}
```

```java
public class ExcelReadTest {

	@Test
	public void read() {
		AnnotationBasedRowSetMapper<RowModel> rowSetMapper
                    = new AnnotationBasedRowSetMapper<>(RowModel.class);
		RowSetReader rowSetReader = RowSetReader.open(file.getInputStream());
		for (RowSet rowSet : rowSetReader) {
			String[] cellValues = rowSet.getRow().getCellValues();
			MappingResult<RowModel> result = rowSetMapper.getMappingResult(rowSet);
			if(!result.hasErrors()){
				RowModel target = result.getTarget();
			}
		}
	}
}
```

## Maven配置

```xml
<dependency>
    <groupId>com.jeeapp</groupId>
    <artifactId>fastexcel</artifactId>
    <version>0.0.1</version>
</dependency>
```

## 联系方式
钉钉群:2880006273
