package com.github.gobars.xlsx;

import lombok.experimental.UtilityClass;
import lombok.val;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.List;

import static java.util.Comparator.reverseOrder;

@UtilityClass
class Poi {
  public void removeRows(Sheet sheet, List<Integer> rows) {
    rows.stream().sorted(reverseOrder()).forEach(r -> removeRow(sheet, r));
  }

  public void removeRow(Sheet sheet, int rowIndex) {
    val lastRowNum = sheet.getLastRowNum();
    if (rowIndex >= 0 && rowIndex < lastRowNum) {
      sheet.shiftRows(rowIndex + 1, lastRowNum, -1);
      return;
    }

    if (rowIndex == lastRowNum) {
      val r = sheet.getRow(rowIndex);
      if (r != null) {
        sheet.removeRow(r);
      }
    }
  }
}
