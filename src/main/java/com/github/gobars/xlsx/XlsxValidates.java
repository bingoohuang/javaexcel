package com.github.gobars.xlsx;

import lombok.experimental.UtilityClass;
import lombok.extern.slf4j.Slf4j;
import lombok.val;

import javax.validation.ConstraintViolation;
import javax.validation.Validation;
import java.lang.reflect.Field;
import java.util.*;

@Slf4j
@UtilityClass
public class XlsxValidates {
  @SuppressWarnings("unchecked")
  public String validate(XlsxOptionTo optionTo, Object obj) {
    if (obj == null) {
      return null;
    }

    if (optionTo != null && optionTo.validatable() != null) {
      return optionTo.validatable().validate(null, obj);
    }

    if (!(obj instanceof Map) && !obj.getClass().isAnnotationPresent(XlsxValid.class)) {
      return null;
    }

    try {
      String msg = javaxValidate(obj);

      if (obj instanceof XlsxValidatable) {
        return ((XlsxValidatable) obj).validate(msg, obj);
      }

      return msg;

    } catch (Exception ex) {
      log.warn("failed to validate obj {}", obj, ex);
    }

    return null;
  }

  public String javaxValidate(Object obj) {
    val v = Validation.buildDefaultValidatorFactory().getValidator();
    Set<ConstraintViolation<Object>> set = v.validate(obj);
    if (set == null || set.isEmpty()) {
      return null;
    }

    val sb = new StringBuilder();
    Set<String> fieldNames = new HashSet<>();

    for (val cv : set) {
      Class cls = cv.getRootBean().getClass();
      val fieldName = cv.getPropertyPath().toString();
      val fields = new ArrayList<>(Arrays.asList(cls.getDeclaredFields()));
      if (cls.getSuperclass() != null) {
        fields.addAll(Arrays.asList(cls.getSuperclass().getDeclaredFields()));
      }

      val name = parseFieldName(fieldName, fields);

      log.warn("validate {}'s value {} failed for {}", name, cv.getInvalidValue(), cv.getMessage());

      if (!fieldNames.contains(fieldName)) {
        fieldNames.add(fieldName);

        sb.append(name).append("格式错误").append(",");
      }
    }

    return sb.substring(0, sb.length() - 1);
  }

  private String parseFieldName(String fieldName, List<Field> fields) {
    for (Field f : fields) {
      if (f.getName().equals(fieldName) && f.isAnnotationPresent(XlsxCol.class)) {
        return XlsxUtil.getTitle(f.getAnnotation(XlsxCol.class));
      }
    }

    return fieldName;
  }
}
