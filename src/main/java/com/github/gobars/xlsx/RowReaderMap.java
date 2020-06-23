package com.github.gobars.xlsx;

import lombok.val;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.HashMap;
import java.util.Map;

class RowReaderMap extends RowReader<Map<String, String>, TitleInfo> {
  public RowReaderMap(Workbook workbook, Map<TitleInfo, FieldInfo> fieldInfos) {
    super(workbook, fieldInfos);
  }

  @Override
  public Map<String, String> newInstance() {
    return new HashMap<>();
  }

  @Override
  public boolean readRow(Map<String, String> map, Row row, OptionTo optionTo) {
    for (val entry : fieldInfos.entrySet()) {
      Cell cell = row.getCell(entry.getValue().index());
      if (cell == null) {
        continue;
      }

      String s = cell.getStringCellValue();

      map.put(entry.getKey().mapKey(), s);
    }

    return true;
  }
}
