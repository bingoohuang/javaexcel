package com.github.gobars.xlsx;

import lombok.experimental.UtilityClass;
import lombok.extern.slf4j.Slf4j;
import lombok.val;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.Closeable;
import java.io.IOException;
import java.util.*;

import static java.util.Comparator.reverseOrder;

@Slf4j
@UtilityClass
public class XlsxUtil {
  public final String EMPTY = "";

  public String trimToEmpty(final String str) {
    return str == null ? EMPTY : str.trim();
  }

  public boolean isEmpty(String s) {
    return s == null || s.length() == 0;
  }

  public boolean isNotEmpty(String s) {
    return s != null && s.length() > 0;
  }

  public void closeQuietly(Closeable closeable) {
    if (closeable != null) {
      try {
        closeable.close();
      } catch (IOException e) {
        // ignore
      }
    }
  }

  public String getTitle(XlsxCol xlsxCol) {
    if (xlsxCol == null) {
      return "";
    }

    if (isNotEmpty(xlsxCol.title())) {
      return xlsxCol.title();
    }

    return xlsxCol.value();
  }

  public boolean contains(String s, String sub) {
    return s != null && s.contains(sub);
  }

  @SuppressWarnings("unchecked")
  public <T> List<T> listOf(T... values) {
    return new ArrayList<>(Arrays.asList(values));
  }

  @SuppressWarnings("unchecked")
  public <T> Map<T, T> mapOf(T... values) {
    HashMap<T, T> m = new HashMap<>(values.length / 2 + 1);

    for (int i = 0; i < values.length; i += 2) {
      m.put(values[i], values[i + 1]);
    }

    return m;
  }

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

  public String getCellValue(Cell cell) {
    if (cell == null) {
      return null;
    }

    switch (cell.getCellType()) {
      case NUMERIC:
        return String.valueOf(cell.getNumericCellValue());
      case BLANK:
        return "";
      case BOOLEAN:
        return String.valueOf(cell.getBooleanCellValue());
      case FORMULA:
        return getFormulaValue(cell);
      case ERROR:
        return FormulaError.forInt(cell.getErrorCellValue()).getString();
      case STRING:
      default:
        return trimToEmpty(cell.getStringCellValue());
    }
  }

  private String getFormulaValue(Cell cell) {
    try {
      return cell.getSheet()
          .getWorkbook()
          .getCreationHelper()
          .createFormulaEvaluator()
          .evaluate(cell)
          .getStringValue();
    } catch (Exception e) {
      log.warn(
          "get formula cell value[{}, {}] error : ", cell.getRowIndex(), cell.getColumnIndex(), e);

      return null;
    }
  }
}
