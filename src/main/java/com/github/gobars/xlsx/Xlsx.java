package com.github.gobars.xlsx;

import lombok.Cleanup;
import lombok.SneakyThrows;
import lombok.val;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.util.*;

public class Xlsx implements Closeable {
  private Workbook workbook;
  private Sheet sheet;

  private Workbook styleWorkbook;
  private Sheet styleSheet;

  /**
   * 指定写入的模板文件.
   *
   * @param fileName 模板文件名
   * @param fileType 模板文件类型
   * @return Xlsx
   */
  public Xlsx template(String fileName, FileType fileType) {
    this.workbook = WorkbookReader.read(fileName, fileType);
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
    if (workbook == null) {
      workbook = new XSSFWorkbook();
    }

    if (beans.isEmpty()) {
      return this;
    }

    val beanClass = beans.get(0).getClass();
    sheet = getSheet();
    val fieldInfos = createFieldFieldInfoMap(beanClass);

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
        Field field = entry.getKey();
        field.setAccessible(true);
        Object fieldValue = field.get(bean);
        if (fieldValue == null) {
          fieldValue = "";
        }

        FieldInfo fi = entry.getValue();
        Cell cell = sheet.getRow(fi.index()).createCell(col);

        cell.setCellValue(fieldValue.toString());
        cell.setCellStyle(fi.dataStyle());
        sheet.setColumnWidth(col, sheet.getColumnWidth(startCol));
      }

      col++;
    }
  }

  private int locateDataColByTitle(Map<Field, FieldInfo> fieldInfos) {
    for (int i = 0, ii = sheet.getLastRowNum(); i <= ii; i++) {
      int titleCol = findAnyTitle(sheet.getRow(i), fieldInfos);
      if (titleCol < 0) {
        continue;
      }

      if (testTitleCol(titleCol, fieldInfos)) {
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

      String cellValue = row.getCell(titleCol).getStringCellValue();

      for (val entry : fieldInfos.entrySet()) {
        if (cellValue.contains(entry.getValue().title())) {
          rowIndexes.put(entry.getKey(), i);
          break;
        }
      }
    }

    if (rowIndexes.size() == fieldInfos.size()) {
      for (val entry : fieldInfos.entrySet()) {
        FieldInfo fi = entry.getValue();
        Integer rowIndex = rowIndexes.get(entry.getKey());
        fi.index(rowIndex);
        fi.dataStyle(sheet.getRow(rowIndex).getCell(titleCol).getCellStyle());
      }
      return true;
    }

    return false;
  }

  private <T> void writeVertical(Map<Field, FieldInfo> fieldInfos, List<T> beans) {
    int startRow = locateDataRowByTitle(fieldInfos);

    Row row = sheet.createRow(startRow);

    if (startRow == 0) {
      writeTileRow(fieldInfos, row);
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

  private <T> int locateDataRowByTitle(Map<T, FieldInfo> fieldInfos) {
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

    val firstFieldInfo = fieldInfos.values().iterator().next();
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
    for (val entry : fieldInfos.entrySet()) {
      int titleColIndex = findTitleInRow(row, entry.getValue().title());
      if (titleColIndex < 0) {
        return false;
      }

      columnIndexes.put(entry.getKey(), titleColIndex);
    }

    for (val entry : fieldInfos.entrySet()) {
      entry.getValue().index(columnIndexes.get(entry.getKey()));
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

  private Map<Field, FieldInfo> createFieldFieldInfoMap(Class<?> beanClass) {
    Map<Field, FieldInfo> fieldInfos = new LinkedHashMap<>();

    for (val field : beanClass.getDeclaredFields()) {
      prepareFieldInfos(fieldInfos, field);
    }
    return fieldInfos;
  }

  private Map<TitleInfo, FieldInfo> createFieldFieldInfoMap(List<TitleInfo> titleInfos) {
    Map<TitleInfo, FieldInfo> fieldInfos = new LinkedHashMap<>();

    for (val field : titleInfos) {
      prepareFieldInfos(fieldInfos, field);
    }
    return fieldInfos;
  }

  private void prepareFieldInfos(Map<Field, FieldInfo> fieldInfos, Field field) {
    val xlsxCol = field.getAnnotation(XlsxCol.class);
    val title = getTitle(xlsxCol);
    if (Util.isEmpty(title)) {
      return;
    }

    FieldInfo fi = new FieldInfo();
    FieldInfo firstFi = getFirstFieldInfo(fieldInfos);
    fi.index(firstFi == null ? 0 : firstFi.index() + fieldInfos.size()).title(title);

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
    field.setAccessible(true);
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

  private Sheet getSheet() {
    if (sheet != null) {
      return sheet;
    }

    if (workbook.getNumberOfSheets() == 0) {
      workbook.createSheet();
    }

    return workbook.getSheetAt(0);
  }

  private String getTitle(XlsxCol xlsxCol) {
    if (xlsxCol == null) {
      return "";
    }

    if (Util.isNotEmpty(xlsxCol.title())) {
      return xlsxCol.title();
    }

    return xlsxCol.value();
  }

  /**
   * 从指定的JavaBean类型，读取JavaBean列表.
   *
   * @param titleInfos 标题信息
   * @return Map列表
   */
  @SneakyThrows
  public List<Map<String, String>> toBeans(List<TitleInfo> titleInfos) {
    ArrayList<Map<String, String>> beans = new ArrayList<>(10);

    if (sheet == null) {
      sheet = getSheet();
    }
    val fieldInfos = createFieldFieldInfoMap(titleInfos);
    int startRow = locateDataRowByTitle(fieldInfos);

    for (int i = startRow, ii = sheet.getLastRowNum(); i <= ii; ++i) {
      beans.add(readMap(fieldInfos, sheet.getRow(i)));
    }

    return beans;
  }

  /**
   * 从指定的JavaBean类型，读取JavaBean列表.
   *
   * @param beanClass JavaBean类型
   * @param <T> JavaBean类型
   * @return JavaBean列表
   */
  @SneakyThrows
  public <T> List<T> toBeans(Class<T> beanClass) {
    ArrayList<T> beans = new ArrayList<>(10);

    if (sheet == null) {
      sheet = getSheet();
    }

    val fieldInfos = createFieldFieldInfoMap(beanClass);
    int startRow = locateDataRowByTitle(fieldInfos);

    for (int i = startRow, ii = sheet.getLastRowNum(); i <= ii; ++i) {
      T t = beanClass.getConstructor().newInstance();
      readRow(t, fieldInfos, sheet.getRow(i));

      beans.add(t);
    }

    return beans;
  }

  @SneakyThrows
  private <T> void readRow(T t, Map<Field, FieldInfo> fieldInfos, Row row) {
    for (val entry : fieldInfos.entrySet()) {
      Cell cell = row.getCell(entry.getValue().index());
      if (cell == null) {
        continue;
      }

      String s = cell.getStringCellValue();

      val field = entry.getKey();
      field.setAccessible(true);
      field.set(t, s);
    }
  }

  @SneakyThrows
  private Map<String, String> readMap(Map<TitleInfo, FieldInfo> fieldInfos, Row row) {
    Map<String, String> map = new HashMap<>(fieldInfos.size());

    for (val entry : fieldInfos.entrySet()) {
      Cell cell = row.getCell(entry.getValue().index());
      if (cell == null) {
        continue;
      }

      val titleInfo = entry.getKey();
      map.put(titleInfo.mapKey(), cell.getStringCellValue());
    }

    return map;
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
   * 写入到Http响应中提供下载.
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

  @SneakyThrows
  public Xlsx read(String fileName, FileType fileType) {
    this.workbook = WorkbookReader.read(fileName, fileType);
    return this;
  }

  @SneakyThrows
  public Xlsx read(String fileName) {
    this.workbook = WorkbookReader.read(fileName, FileType.NORMAL);
    return this;
  }

  @SneakyThrows
  public Xlsx read(InputStream is) {
    this.workbook = WorkbookReader.read(is);
    return this;
  }

  @Override
  public void close() {
    if (this.workbook != null) {
      closeQuietly(this.workbook);
    }
  }

  private void closeQuietly(Closeable closeable) {
    try {
      closeable.close();
    } catch (IOException e) {
      // ignore
    }
  }
}
