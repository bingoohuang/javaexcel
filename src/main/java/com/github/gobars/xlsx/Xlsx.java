package com.github.gobars.xlsx;

import lombok.Cleanup;
import lombok.SneakyThrows;
import lombok.val;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.Closeable;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class Xlsx implements Closeable {
  boolean autoClose = true;
  private Workbook workbook;
  private Sheet sheet;
  private Workbook styleWorkbook;
  private Sheet styleSheet;

  /**
   * 是否在readBeans或者write方法之后，自动关闭相关资源.
   *
   * @param autoClose 是否自动关闭.
   * @return Xlsx
   */
  public Xlsx autoClose(boolean autoClose) {
    this.autoClose = autoClose;
    return this;
  }

  /**
   * 指定样式文件.
   *
   * @param fileName 样式文件名
   * @param fileType 样式文件类型
   * @return Xlsx
   */
  public Xlsx style(String fileName, FileType fileType) {
    this.styleWorkbook = WorkbookReader.read(fileName, fileType);
    return this;
  }

  /**
   * 写入列表.
   *
   * @param beans bean列表
   * @param <T> bean类型
   * @param fromOptions 写入选项
   * @return Xlsx
   */
  @SneakyThrows
  public <T> Xlsx fromBeans(List<T> beans, FromOption... fromOptions) {
    getSheet();

    val beanClass = beans.get(0).getClass();
    val fieldInfos = createFieldInfoMap(beanClass);

    val option = fromOptions.length > 0 ? fromOptions[0] : new FromOption();
    if (option.horizontal()) {
      writeHorizontal(fieldInfos, beans);
    } else {
      writeVertical(fieldInfos, beans);
    }

    return this;
  }

  @SneakyThrows
  private <T> void writeHorizontal(Map<Field, FieldInfo> fieldInfos, List<T> beans) {
    int startCol = locateDataColByTitle(fieldInfos);
    int col = startCol;

    for (val bean : beans) {
      for (val entry : fieldInfos.entrySet()) {
        Object fieldValue = entry.getKey().get(bean);
        if (fieldValue == null) {
          fieldValue = "";
        }

        FieldInfo fi = entry.getValue();
        Cell cell = sheet.getRow(fi.index()).createCell(col);

        cell.setCellValue(fieldValue.toString());
        cell.setCellStyle(fi.dataStyle());
      }

      sheet.setColumnWidth(col, sheet.getColumnWidth(startCol));
      col++;
    }
  }

  private int locateDataColByTitle(Map<Field, FieldInfo> fieldInfos) {
    for (int i = 0, ii = sheet.getLastRowNum(); i <= ii; i++) {
      int titleCol = findAnyTitle(sheet.getRow(i), fieldInfos);
      if (titleCol >= 0 && testTitleCol(titleCol, fieldInfos)) {
        return titleCol;
      }
    }

    return 0;
  }

  private boolean testTitleCol(int titleCol, Map<Field, FieldInfo> fieldInfos) {
    Map<Field, Integer> rowIndexes = new HashMap<>(fieldInfos.size());

    for (int i = 0, ii = sheet.getLastRowNum(); i <= ii; i++) {
      Row row = sheet.getRow(i);
      if (row == null || row.getCell(titleCol) == null) {
        continue;
      }

      String title = row.getCell(titleCol).getStringCellValue();
      for (val entry : fieldInfos.entrySet()) {
        if (title.contains(entry.getValue().title())) {
          rowIndexes.put(entry.getKey(), i);
          break;
        }
      }
    }

    if (rowIndexes.size() == fieldInfos.size()) {
      for (val entry : fieldInfos.entrySet()) {
        int rownum = rowIndexes.get(entry.getKey());

        FieldInfo fi = entry.getValue();
        fi.index(rownum);
        fi.dataStyle(sheet.getRow(rownum).getCell(titleCol).getCellStyle());
      }

      return true;
    }

    return false;
  }

  private <T> void writeVertical(Map<Field, FieldInfo> fieldInfos, List<T> beans) {
    int startRow = locateDataRowByTitle(fieldInfos);

    if (startRow == 0) {
      writeTileRow(fieldInfos, sheet.createRow(startRow));
      startRow = 1;
    }

    for (val bean : beans) {
      writeDataRow(fieldInfos, startRow++, bean);
    }
  }

  private <T> void writeDataRow(Map<Field, FieldInfo> fieldInfos, int rowIndex, T bean) {
    val row = sheet.createRow(rowIndex);

    for (val entry : fieldInfos.entrySet()) {
      writeCell(row, bean, entry.getValue(), entry.getKey());
    }
  }

  private void writeTileRow(Map<Field, FieldInfo> fieldInfos, Row row) {
    for (val fieldInfo : fieldInfos.entrySet()) {
      val fi = fieldInfo.getValue();
      val cell = row.createCell(fi.index());
      cell.setCellValue(fi.title());
      if (fi.titleStyle() != null) {
        cell.setCellStyle(fi.titleStyle());
      }
    }
  }

  <T> int locateDataRowByTitle(Map<T, FieldInfo> fieldInfos) {
    for (int i = 0, ii = sheet.getLastRowNum(); i <= ii; i++) {
      Row row = sheet.getRow(i);
      if (row != null && findAllTitles(row, fieldInfos)) {
        fulfilDataCellStyle(fieldInfos, i + 1);

        return i + 1;
      }
    }

    return 0;
  }

  private <T> void fulfilDataCellStyle(Map<T, FieldInfo> fieldInfos, int dataRowNum) {
    Row dataRow = sheet.getRow(dataRowNum);
    if (dataRow == null) {
      return;
    }

    val firstFieldInfo = getFirstFieldInfo(fieldInfos);
    Cell dataCell = dataRow.getCell(firstFieldInfo.index());
    if (dataCell == null) {
      return;
    }

    CellStyle cellStyle = dataCell.getCellStyle();

    for (val entry : fieldInfos.entrySet()) {
      entry.getValue().dataStyle(cellStyle);
    }
  }

  private int findAnyTitle(Row row, Map<Field, FieldInfo> fieldInfos) {
    if (row == null) {
      return -1;
    }

    for (val entry : fieldInfos.entrySet()) {
      int titleColIndex = findTitleInRow(row, entry.getValue().title());
      if (titleColIndex >= 0) {
        return titleColIndex;
      }
    }

    return -1;
  }

  private <T> boolean findAllTitles(Row row, Map<T, FieldInfo> fieldInfos) {
    Map<T, Integer> columnIndexes = new HashMap<>(fieldInfos.size());
    for (val i : fieldInfos.entrySet()) {
      int titleColIndex = findTitleInRow(row, i.getValue().title());
      if (titleColIndex < 0) {
        return false;
      }

      columnIndexes.put(i.getKey(), titleColIndex);
    }

    for (val i : fieldInfos.entrySet()) {
      i.getValue().index(columnIndexes.get(i.getKey()));
    }

    return true;
  }

  private int findTitleInRow(Row row, String title) {
    for (int i = 0, ii = row.getLastCellNum(); i < ii; i++) {
      Cell cell = row.getCell(i);
      if (cell != null && cell.getStringCellValue().contains(title)) {
        return i;
      }
    }

    return -1;
  }

  private Map<Field, FieldInfo> createFieldInfoMap(Class<?> beanClass) {
    Map<Field, FieldInfo> fieldInfos = new LinkedHashMap<>();

    for (val f : beanClass.getDeclaredFields()) {
      f.setAccessible(true);
      prepareFieldInfos(fieldInfos, f);
    }

    return fieldInfos;
  }

  private Map<TitleInfo, FieldInfo> createFieldInfoMap(List<TitleInfo> titleInfos) {
    Map<TitleInfo, FieldInfo> fieldInfos = new LinkedHashMap<>();

    for (val f : titleInfos) {
      prepareFieldInfos(fieldInfos, f);
    }

    return fieldInfos;
  }

  private void prepareFieldInfos(Map<Field, FieldInfo> fieldInfos, Field field) {
    val xlsxCol = field.getAnnotation(XlsxCol.class);
    val title = Util.getTitle(xlsxCol);
    if (Util.isEmpty(title)) {
      return;
    }

    FieldInfo fi = new FieldInfo();
    FieldInfo firstFi = getFirstFieldInfo(fieldInfos);
    fi.index(firstFi == null ? 0 : firstFi.index() + fieldInfos.size())
        .title(title)
        .ignoreRow(xlsxCol.ignoreRow());

    fieldInfos.put(field, fi);

    if (styleWorkbook == null) {
      return;
    }

    String titleStyle = xlsxCol.titleStyle();
    if (Util.isNotEmpty(titleStyle)) {
      fi.titleStyle(cloneCellStyle(titleStyle));
    } else if (firstFi != null) {
      // 继承第一个注解的样式
      fi.titleStyle(firstFi.titleStyle());
    }

    String dataStyle = xlsxCol.dataStyle();
    if (Util.isNotEmpty(dataStyle)) {
      fi.dataStyle(cloneCellStyle(dataStyle));
    } else if (firstFi != null) {
      // 继承第一个注解的样式
      fi.dataStyle(firstFi.dataStyle());
    }
  }

  private void prepareFieldInfos(Map<TitleInfo, FieldInfo> fieldInfos, TitleInfo field) {
    FieldInfo fi = new FieldInfo();
    FieldInfo firstFi = getFirstFieldInfo(fieldInfos);
    fi.index(firstFi == null ? 0 : firstFi.index() + fieldInfos.size()).title(field.title());

    fieldInfos.put(field, fi);
  }

  private <T> FieldInfo getFirstFieldInfo(Map<T, FieldInfo> fieldInfos) {
    if (fieldInfos.isEmpty()) {
      return null;
    }

    return fieldInfos.values().iterator().next();
  }

  @SneakyThrows
  private <T> void writeCell(Row row, T bean, FieldInfo fi, Field field) {
    Object fieldValue = field.get(bean);
    if (fieldValue == null) {
      fieldValue = "";
    }

    Cell cell = row.createCell(fi.index());
    cell.setCellValue(fieldValue.toString());

    if (fi.dataStyle() != null) {
      cell.setCellStyle(fi.dataStyle());
    }
  }

  private CellStyle cloneCellStyle(String cellReference) {
    if (styleSheet == null) {
      styleSheet = styleWorkbook.getSheetAt(0);
    }

    val cr = new CellReference(cellReference);
    val style = styleSheet.getRow(cr.getRow()).getCell(cr.getCol()).getCellStyle();
    val cloneStyle = workbook.createCellStyle();
    cloneStyle.cloneStyleFrom(style);
    return cloneStyle;
  }

  Sheet getSheet() {
    if (workbook == null) {
      workbook = new XSSFWorkbook();
    }

    if (sheet == null) {
      if (workbook.getNumberOfSheets() == 0) {
        sheet = workbook.createSheet();
      } else {
        sheet = workbook.getSheetAt(0);
      }
    }

    return sheet;
  }

  /**
   * 从指定的JavaBean类型，读取JavaBean列表.
   *
   * @param titleInfos 标题信息
   * @return Map列表
   */
  public List<Map<String, String>> toBeans(List<TitleInfo> titleInfos, ToOption... toOptions) {
    val fieldInfos = createFieldInfoMap(titleInfos);

    return new MapRowToBeans(workbook).toBeans(this, fieldInfos, null, toOptions);
  }

  /**
   * 从指定的JavaBean类型，读取JavaBean列表.
   *
   * @param beanClass JavaBean类型
   * @param <T> JavaBean类型
   * @param toOptions 写入选项
   * @return JavaBean列表
   */
  @SneakyThrows
  public <T> List<T> toBeans(Class<T> beanClass, ToOption... toOptions) {
    val fieldInfos = createFieldInfoMap(beanClass);

    XlsxValid xv = beanClass.getAnnotation(XlsxValid.class);
    return new BeanRowToBeans<T>(workbook, beanClass).toBeans(this, fieldInfos, xv, toOptions);
  }

  void doAutoClose() {
    if (this.autoClose) {
      close();
    }
  }

  /**
   * 用密码保护工作簿。只有在xlsx格式才起作用。
   *
   * @param password 保护密码。
   * @return Xlsx
   */
  public Xlsx protect(String password) {
    if (Util.isEmpty(password)) {
      throw new XlsxException("password should not be empty");
    }

    if (workbook instanceof XSSFWorkbook) {
      val xw = (XSSFWorkbook) workbook;
      for (int i = 0, ii = xw.getNumberOfSheets(); i < ii; ++i) {
        xw.getSheetAt(i).protectSheet(password);
      }

      return this;
    }

    throw new XlsxException("only xlsx supported");
  }

  /**
   * 写入到输出流中.
   *
   * @param out OutputStream
   * @return Xlsx
   */
  @SneakyThrows
  public Xlsx write(OutputStream out) {
    workbook.write(out);
    doAutoClose();

    return this;
  }

  /**
   * 写入到文件中.
   *
   * @param fileName 文件名
   * @return Xlsx
   */
  @SneakyThrows
  public Xlsx write(String fileName) {
    @Cleanup val out = new FileOutputStream(fileName);
    return write(out);
  }

  /**
   * 写入到Http响应流中提供下载.
   *
   * <p>注意：调用完此方法后，请不要再往输出流中写入其它数据，否则会导致excel下载文件错误.
   *
   * @param fileName 下载文件名
   * @param r HTTP响应流
   * @return Xlsx
   */
  @SneakyThrows
  public Xlsx write(String fileName, HttpServletResponse r) {
    r.setContentType("application/vnd.ms-excel;charset=UTF-8");
    val f = URLEncoder.encode(fileName, "UTF-8");
    val v = "attachment; filename=\"" + f + "\"; filename*=utf-8'zh_cn'" + f;
    r.setHeader("Content-disposition", v);
    @Cleanup val out = r.getOutputStream();

    return write(out);
  }

  /**
   * 指定写入的模板文件，或者需要读取的文件.
   *
   * @param fileName 文件名
   * @param fileType 文件类型
   * @return Xlsx
   */
  public Xlsx read(String fileName, FileType fileType) {
    this.workbook = WorkbookReader.read(fileName, fileType);
    return this;
  }

  /**
   * 从本地文件指定写入的模板文件，或者需要读取的文件.
   *
   * @param fileName 文件名
   * @return Xlsx
   */
  public Xlsx read(String fileName) {
    return read(fileName, FileType.NORMAL);
  }

  /**
   * 从输入流指定写入的模板文件，或者需要读取的文件.
   *
   * @param is 输入流
   * @return Xlsx
   */
  public Xlsx read(InputStream is) {
    this.workbook = WorkbookReader.read(is);
    return this;
  }

  @Override
  public void close() {
    Util.closeQuietly(this.workbook);
    Util.closeQuietly(this.styleWorkbook);
    this.workbook = null;
    this.styleWorkbook = null;
    this.sheet = null;
    this.styleSheet = null;
  }
}
