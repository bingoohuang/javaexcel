package com.github.gobars.xlsx;

import lombok.SneakyThrows;
import lombok.val;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Field;
import java.util.Map;

class RowReaderBean<T> extends RowReader<T, Field> {
  private final Class<T> beanClass;

  public RowReaderBean(Workbook workbook, Map<Field, FieldInfo> fieldInfos, Class<T> beanClass) {
    super(workbook, fieldInfos);
    this.beanClass = beanClass;
  }

  @Override
  @SneakyThrows
  public T newInstance() {
    return beanClass.getConstructor().newInstance();
  }

  @Override
  @SneakyThrows
  public boolean readRow(T t, Row row, OptionTo optionTo) {
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
