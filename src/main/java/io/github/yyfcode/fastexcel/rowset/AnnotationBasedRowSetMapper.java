package io.github.yyfcode.fastexcel.rowset;

import java.lang.reflect.Field;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

import org.apache.commons.lang3.reflect.FieldUtils;
import org.springframework.context.support.EmbeddedValueResolutionSupport;
import org.springframework.format.AnnotationFormatterFactory;
import org.springframework.format.Formatter;
import org.springframework.format.Parser;
import org.springframework.format.Printer;
import org.springframework.format.datetime.DateFormatter;
import org.springframework.format.number.NumberStyleFormatter;
import org.springframework.format.support.DefaultFormattingConversionService;
import org.springframework.util.NumberUtils;
import org.springframework.util.StringUtils;
import io.github.yyfcode.fastexcel.annotation.ExcelProperty;

/**
 * @author Justice
 */
public class AnnotationBasedRowSetMapper<T> extends BeanWrapperRowSetMapper<T> {

	private final Class<? extends T> type;

	public AnnotationBasedRowSetMapper(Class<? extends T> type) {
		super(type);
		this.type = type;
		DefaultFormattingConversionService conversionService = new DefaultFormattingConversionService();
		conversionService.addFormatterForFieldAnnotation(new ExcelPropertyAnnotationFormatterFactory());
		this.setConversionService(conversionService);
	}

	@Override
	public Map<String, Integer> getFieldColumns(RowSet rowSet) {
		Map<String, Integer> fieldColumns = super.getFieldColumns(rowSet);
		if (fieldColumns == null) {
			fieldColumns = FieldUtils.getFieldsListWithAnnotation(type, ExcelProperty.class)
				.stream()
				.collect(Collectors.toMap(Field::getName, this::getColumnIndex));
			setFieldColumns(fieldColumns);
		}
		return fieldColumns;
	}

	private Integer getColumnIndex(Field field) {
		return field.getAnnotation(ExcelProperty.class).column();
	}

	static class ExcelPropertyAnnotationFormatterFactory extends EmbeddedValueResolutionSupport implements AnnotationFormatterFactory<ExcelProperty> {

		private static final Set<Class<?>> FIELD_TYPES;

		private static final Set<Class<?>> DATE_TIME_FIELD_TYPES;

		static {
			Set<Class<?>> dateTimeFieldTypes = new HashSet<>();
			dateTimeFieldTypes.add(Date.class);
			dateTimeFieldTypes.add(Calendar.class);
			dateTimeFieldTypes.add(Long.class);
			DATE_TIME_FIELD_TYPES = Collections.unmodifiableSet(dateTimeFieldTypes);
			Set<Class<?>> fieldTypes = new HashSet<>();
			fieldTypes.addAll(DATE_TIME_FIELD_TYPES);
			fieldTypes.addAll(NumberUtils.STANDARD_NUMBER_TYPES);
			FIELD_TYPES = Collections.unmodifiableSet(fieldTypes);
		}

		@Override
		public Set<Class<?>> getFieldTypes() {
			return FIELD_TYPES;
		}

		@Override
		public Printer<?> getPrinter(ExcelProperty annotation, Class<?> fieldType) {
			if (DATE_TIME_FIELD_TYPES.contains(fieldType)) {
				return getDateFormatter(annotation);
			} else {
				return getNumberFormatter(annotation);
			}
		}

		@Override
		public Parser<?> getParser(ExcelProperty annotation, Class<?> fieldType) {
			if (DATE_TIME_FIELD_TYPES.contains(fieldType)) {
				return getDateFormatter(annotation);
			} else {
				return getNumberFormatter(annotation);
			}
		}

		protected Formatter<Number> getNumberFormatter(ExcelProperty annotation) {
			String pattern = resolveEmbeddedValue(annotation.format());
			if (StringUtils.hasLength(pattern)) {
				return new NumberStyleFormatter(pattern);
			} else {
				return new NumberStyleFormatter();
			}
		}

		protected Formatter<Date> getDateFormatter(ExcelProperty annotation) {
			DateFormatter formatter = new DateFormatter();
			String pattern = resolveEmbeddedValue(annotation.format());
			if (StringUtils.hasLength(pattern)) {
				formatter.setPattern(pattern);
			}
			return formatter;
		}
	}
}
