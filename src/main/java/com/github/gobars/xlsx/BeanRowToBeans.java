package com.github.gobars.xlsx;

import lombok.SneakyThrows;
import lombok.val;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Field;
import java.util.Map;

public class BeanRowToBeans<T> extends AbstractRowToBeans<T, Field> {
  private final Class<T> beanClass;

  public BeanRowToBeans(Workbook workbook, Class<T> beanClass) {
    super(workbook);
    this.beanClass = beanClass;
  }

  @Override
  @SneakyThrows
  public T newInstance() {
    return beanClass.getConstructor().newInstance();
  }

  @Override
  @SneakyThrows
  public boolean readRow(T t, Map<Field, FieldInfo> fieldInfos, Row row, ToOption toOption) {
    for (val entry : fieldInfos.entrySet()) {
      Cell cell = row.getCell(entry.getValue().index());
      if (cell == null) {
        continue;
      }

      String s = cell.getStringCellValue();
      String ignoreRow = entry.getValue().ignoreRow();
      if (Util.isNotEmpty(ignoreRow) && Util.contains(s, ignoreRow)) {
        return false;
      }

      entry.getKey().set(t, s);
    }

    return true;
  }
}
