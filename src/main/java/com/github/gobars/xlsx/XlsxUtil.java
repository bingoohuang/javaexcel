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

  public void closeQuietly(Object closeable) {
    if (closeable instanceof Closeable) {
      try {
        ((Closeable) closeable).close();
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
    return new ArrayList<T>(Arrays.asList(values));
  }

  @SuppressWarnings("unchecked")
  public <T> Map<T, T> mapOf(T... values) {
    HashMap<T, T> m = new HashMap<T, T>(values.length / 2 + 1);

    for (int i = 0; i < values.length; i += 2) {
      m.put(values[i], values[i + 1]);
    }

    return m;
  }

  public void removeRows(Sheet sheet, List<Integer> rows) {
    Collections.sort(rows, Collections.reverseOrder());
    for (val r : rows) {
      removeRow(sheet, r);
    }
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
      case Cell.CELL_TYPE_NUMERIC:
        return fmt(cell.getNumericCellValue());
      case Cell.CELL_TYPE_BLANK:
        return "";
      case Cell.CELL_TYPE_BOOLEAN:
        return String.valueOf(cell.getBooleanCellValue());
      case Cell.CELL_TYPE_FORMULA:
        return getFormulaValue(cell);
      case Cell.CELL_TYPE_ERROR:
        return FormulaError.forInt(cell.getErrorCellValue()).getString();
      case Cell.CELL_TYPE_STRING:
      default:
        return trimToEmpty(cell.getStringCellValue());
    }
  }

  public String fmt(double d) {
    return d == (long) d ? String.format("%d", (long) d) : String.format("%s", d);
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

  /**
   * Test if s is any of element in the values.
   *
   * @param s tested one.
   * @param values list of values
   * @param <T> s type.
   * @return true if s is one of the values.
   */
  public <T> boolean anyOf(T s, T... values) {
    for (val i : values) {
      if (s.equals(i)) return true;
    }

    return false;
  }
}
