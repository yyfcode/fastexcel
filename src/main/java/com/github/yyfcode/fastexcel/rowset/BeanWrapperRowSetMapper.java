package com.github.yyfcode.fastexcel.rowset;

import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Properties;

import org.springframework.beans.BeanUtils;
import org.springframework.beans.MutablePropertyValues;
import org.springframework.context.MessageSource;
import org.springframework.context.support.MessageSourceAccessor;
import org.springframework.context.support.ResourceBundleMessageSource;
import org.springframework.core.convert.ConversionService;
import org.springframework.util.Assert;
import org.springframework.validation.BindingResult;
import org.springframework.validation.DataBinder;
import org.springframework.validation.FieldError;
import org.springframework.validation.Validator;

/**
 * @author Justice
 */
public class BeanWrapperRowSetMapper<T> implements RowSetMapper<T> {

	private final Class<? extends T> type;

	private boolean strict = true;

	private Validator validator;

	private MessageSourceAccessor messages = DefaultMessageSource.getAccessor();

	private ConversionService conversionService;

	private Map<String, Integer> fieldColumns;

	public BeanWrapperRowSetMapper(Class<? extends T> type) {
		this.type = type;
	}

	@Override
	public T mapRowSet(RowSet rowSet) throws MappingException {
		T copy = BeanUtils.instantiateClass(type);
		DataBinder binder = createBinder(copy);
		binder.bind(new MutablePropertyValues(getProperties(rowSet)));
		binder.validate();
		BindingResult bindingResult = binder.getBindingResult();
		if (bindingResult.hasErrors()) {
			Map<String, Integer> fieldColumns = getFieldColumns(rowSet);
			List<FieldError> fieldErrors = bindingResult.getFieldErrors();
			MappingException ex = new MappingException(rowSet.getRow());
			for (FieldError fieldError : fieldErrors) {
				ex.addError(fieldColumns.get(fieldError.getField()), messages.getMessage(fieldError));
			}
			throw ex;
		}
		return copy;
	}


	protected Properties getProperties(RowSet rowSet) {
		Map<String, Integer> fieldMappings = getFieldColumns(rowSet);
		if (fieldMappings == null) {
			throw new IllegalStateException("Cannot create properties without meta data");
		}
		String[] values = rowSet.getRow().getCellValues();
		int length = values.length;
		Properties props = new Properties();
		for (Entry<String, Integer> fieldMapping : fieldMappings.entrySet()) {
			if (fieldMapping.getValue() < length) {
				String value = values[fieldMapping.getValue()];
				if (value != null) {
					props.setProperty(fieldMapping.getKey(), value);
				}
			}
		}
		return props;
	}

	@Override
	public MappingResult<T> getMappingResult(RowSet rowSet) {
		MappingResult<T> mappingResult = RowSetMapper.super.getMappingResult(rowSet);
		mappingResult.setFieldColumns(fieldColumns);
		return mappingResult;
	}

	protected Map<String, Integer> getFieldColumns(RowSet rowSet) {
		return fieldColumns;
	}

	protected DataBinder createBinder(Object target) {
		DataBinder binder = new DataBinder(target);
		binder.setIgnoreUnknownFields(!this.strict);
		initBinder(binder);
		if (this.conversionService != null) {
			binder.setConversionService(this.conversionService);
		}
		if (this.validator != null) {
			binder.setValidator(this.validator);
		}
		return binder;
	}

	protected void initBinder(DataBinder binder) {
	}

	public void setStrict(boolean strict) {
		this.strict = strict;
	}

	public void setValidator(Validator validator) {
		this.validator = validator;
	}

	public void setMessageSource(MessageSource messageSource) {
		Assert.notNull(messageSource, "messageSource cannot be null");
		this.messages = new MessageSourceAccessor(messageSource);
	}

	public void setConversionService(ConversionService conversionService) {
		this.conversionService = conversionService;
	}

	public void setFieldColumns(Map<String, Integer> fieldColumns) {
		this.fieldColumns = fieldColumns;
	}

	static class DefaultMessageSource extends ResourceBundleMessageSource {

		public DefaultMessageSource() {
			setBasename("com.platform.excel.messages");
		}

		public static MessageSourceAccessor getAccessor() {
			return new MessageSourceAccessor(new DefaultMessageSource());
		}
	}
}
