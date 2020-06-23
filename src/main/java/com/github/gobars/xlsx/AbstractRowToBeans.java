package com.github.gobars.xlsx;

import lombok.val;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * 读取Excel行到JavaBean或者Map
 *
 * @param <T> JavaBean的Class，或者Map
 * @param <K> TitleInfo或者Field
 * @author bingoo
 */
public abstract class AbstractRowToBeans<T, K> {
  private final Workbook workbook;
  private CellStyle redCellStyle;

  public AbstractRowToBeans(Workbook workbook) {
    this.workbook = workbook;
  }

  public abstract T newInstance();

  public abstract boolean readRow(T t, Map<K, FieldInfo> fieldInfos, Row row, ToOption toOption);

  public List<T> toBeans(Xlsx x, Map<K, FieldInfo> fieldInfos, XlsxValid xv, ToOption[] toOptions) {
    ArrayList<T> beans = new ArrayList<>(10);

    val sh = x.getSheet();
    int startRow = x.locateDataRowByTitle(fieldInfos);
    int errColNum = sh.getRow(startRow).getLastCellNum();

    val okRows = new ArrayList<Integer>();
    ToOption toOption = createToOption(toOptions);
    int errRows = 0;

    boolean tempAutoClose = x.autoClose;

    for (int i = startRow, ii = sh.getLastRowNum(); i <= ii; ++i) {
      T t = newInstance();
      Row row = sh.getRow(i);

      if (readRow(t, fieldInfos, row, toOption) && !shouldIgnore(toOption, t)) {
        String errMsg = ValidateUtil.validate(toOption, t);
        if (errMsg == null) {
          okRows.add(i);
          beans.add(t);
        } else if (writeErrorToExcel(xv, toOption)) {
          errRows++;
          Cell errCell = row.createCell(errColNum);
          errCell.setCellValue(errMsg);
          errCell.setCellStyle(createRedCellStyle());
          x.autoClose(false);
          errCallback(toOption, t, errMsg, i);
        }
      }
    }

    if (removeOkRows(xv, okRows, toOption)) {
      Poi.removeRows(sh, okRows);
      x.autoClose(false);
    }

    toOption.errRows(errRows).okRows(okRows.size());

    x.doAutoClose();

    if (tempAutoClose != x.autoClose) {
      x.autoClose = tempAutoClose;
    }

    return beans;
  }

  @SuppressWarnings("unchecked")
  private boolean shouldIgnore(ToOption toOption, T map) {
    val cb = toOption.ignoreCallback();

    return cb != null && cb.shouldIgnore(map);
  }

  private boolean removeOkRows(XlsxValid xv, ArrayList<Integer> okRows, ToOption toOption) {
    return !okRows.isEmpty() && (toOption.removeOkRows() || xv != null && xv.removeOKRows());
  }

  private boolean writeErrorToExcel(XlsxValid xv, ToOption toOption) {
    return toOption.writeErrorToExcel() || xv != null && xv.writeErrorToExcel();
  }

  @SuppressWarnings("unchecked")
  private void errCallback(ToOption toOption, T t, String errMsg, int rownum) {
    if (toOption.errCallback() != null) {
      toOption.errCallback().call(t, errMsg, rownum);
    }
  }

  private ToOption createToOption(ToOption[] toOptions) {
    if (toOptions.length == 0) {
      return new ToOption();
    }

    return toOptions[0];
  }

  private CellStyle createRedCellStyle() {
    if (redCellStyle != null) {
      return redCellStyle;
    }

    redCellStyle = workbook.createCellStyle();
    val font = workbook.createFont();
    font.setColor(IndexedColors.RED.getIndex());
    redCellStyle.setFont(font);

    return redCellStyle;
  }
}
