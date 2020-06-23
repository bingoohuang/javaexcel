package com.github.gobars.xlsx;

import lombok.val;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.HashMap;
import java.util.Map;

public class MapRowToBeans extends AbstractRowToBeans<Map<String, String>, TitleInfo> {
  public MapRowToBeans(Workbook workbook) {
    super(workbook);
  }

  @Override
  public Map<String, String> newInstance() {
    return new HashMap<>();
  }

  @Override
  public boolean readRow(
      Map<String, String> map, Map<TitleInfo, FieldInfo> fieldInfos, Row row, ToOption toOption) {
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
