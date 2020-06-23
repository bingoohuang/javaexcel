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
class RowReader<T, K> {
  final Workbook workbook;
  final Map<K, FieldInfo> fieldInfos;
  private CellStyle redCellStyle;

  public RowReader(Workbook workbook, Map<K, FieldInfo> fieldInfos) {
    this.workbook = workbook;
    this.fieldInfos = fieldInfos;
  }

  public T newInstance() {
    return null;
  }

  public boolean readRow(T t, Row row, OptionTo optionTo) {
    return false;
  }

  public List<T> toBeans(Xlsx x, XlsxValid xv, OptionTo[] optionTos) {
    ArrayList<T> beans = new ArrayList<>(10);

    val sh = x.getSheet();
    int startRow = x.locateDataRowByTitle(fieldInfos);
    int errColNum = sh.getRow(startRow).getLastCellNum();

    val okRows = new ArrayList<Integer>();
    OptionTo optionTo = createToOption(optionTos);
    int errRows = 0;

    boolean tempAutoClose = x.autoClose;

    for (int i = startRow, ii = sh.getLastRowNum(); i <= ii; ++i) {
      T t = newInstance();
      Row row = sh.getRow(i);

      if (readRow(t, row, optionTo) && !shouldIgnore(optionTo, t)) {
        String errMsg = ValidateUtil.validate(optionTo, t);
        if (errMsg == null) {
          okRows.add(i);
          beans.add(t);
        } else if (writeErrorToExcel(xv, optionTo)) {
          errRows++;
          Cell errCell = row.createCell(errColNum);
          errCell.setCellValue(errMsg);
          errCell.setCellStyle(createRedCellStyle());
          x.autoClose(false);
          errCallback(optionTo, t, errMsg, i);
        }
      }
    }

    if (removeOkRows(xv, okRows, optionTo)) {
      Poi.removeRows(sh, okRows);
      x.autoClose(false);
    }

    optionTo.errRows(errRows).okRows(okRows.size());

    x.doAutoClose();

    if (tempAutoClose != x.autoClose) {
      x.autoClose = tempAutoClose;
    }

    return beans;
  }

  @SuppressWarnings("unchecked")
  private boolean shouldIgnore(OptionTo optionTo, T t) {
    val cb = optionTo.ignoreCallback();

    return cb != null && cb.shouldIgnore(t);
  }

  private boolean removeOkRows(XlsxValid xv, ArrayList<Integer> okRows, OptionTo optionTo) {
    return !okRows.isEmpty() && (optionTo.removeOkRows() || xv != null && xv.removeOKRows());
  }

  private boolean writeErrorToExcel(XlsxValid xv, OptionTo optionTo) {
    return optionTo.writeErrorToExcel() || xv != null && xv.writeErrorToExcel();
  }

  @SuppressWarnings("unchecked")
  private void errCallback(OptionTo optionTo, T t, String errMsg, int rownum) {
    if (optionTo.errCallback() != null) {
      optionTo.errCallback().call(t, errMsg, rownum);
    }
  }

  private OptionTo createToOption(OptionTo[] optionTos) {
    if (optionTos.length == 0) {
      return new OptionTo();
    }

    return optionTos[0];
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
