# fastexcel

![xxx](https://img.shields.io/badge/version-1.0.0-green) ![xxx](https://img.shields.io/badge/jdk-1.8-green) ![xxx](https://img.shields.io/badge/poi-5.0.0-green) ![xxx](https://img.shields.io/badge/spring-5.3.23-green)

A Java library for reading and writing big excel files quickly and easily.
In this file you can find some simple examples.
You can also take a look on a working Spring Boot app that uses this library: [https://github.com/yyfcode/fastexcel-demo](https://github.com/yyfcode/fastexcel-demo)

See the [中文文档](https://github.com/yyfcode/fastexcel/blob/master/README-zh.md)  for Chinese readme.

## Examples

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

## Maven configuration

```xml
<dependency>
    <groupId>com.jeeapp</groupId>
    <artifactId>fastexcel</artifactId>
    <version>0.0.1</version>
</dependency>
```
