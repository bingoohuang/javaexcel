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
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class Xlsx implements Closeable {
  private Workbook workbook;
  private Sheet sheet;

  private Workbook styleWorkbook;
  private Sheet styleSheet;

  /**
   * 指定模板文件.
   *
   * @param fileName 模板文件名
   * @param fileType 模板文件类型
   * @return Xlsx
   */
  public Xlsx template(String fileName, FileType fileType) {
    this.styleWorkbook = WorkbookReader.read(fileName, fileType);
    return this;
  }

  /**
   * 写入列表.
   *
   * @param beans bean列表
   * @return Xlsx
   */
  @SneakyThrows
  public <T> Xlsx fromBeans(List<T> beans) {
    if (beans.isEmpty()) {
      return this;
    }

    val beanClass = beans.get(0).getClass();

    if (workbook == null) {
      workbook = new XSSFWorkbook();
      sheet = workbook.createSheet();
    } else {
      sheet = getSheet(beanClass);
    }

    Map<Field, FieldInfo> fieldInfos = createFieldFieldInfoMap(beanClass);

    Row row = sheet.createRow(0);
    for (val fieldInfo : fieldInfos.entrySet()) {
      FieldInfo fi = fieldInfo.getValue();
      Cell cell = row.createCell(fi.columnIndex());
      cell.setCellValue(getTitle(fi.xlsxCol()));
      if (fi.titleStyle() != null) {
        cell.setCellStyle(fi.titleStyle());
      }
    }

    for (val bean : beans) {
      row = sheet.createRow(sheet.getLastRowNum() + 1);

      for (val entry : fieldInfos.entrySet()) {
        writeCellValue(row, bean, entry.getValue(), entry.getKey());
      }
    }

    return this;
  }

  private Map<Field, FieldInfo> createFieldFieldInfoMap(Class<?> beanClass) {
    Map<Field, FieldInfo> fieldInfos = new LinkedHashMap<>();

    for (val field : beanClass.getDeclaredFields()) {
      prepareFieldInfos(fieldInfos, field);
    }
    return fieldInfos;
  }

  private void prepareFieldInfos(Map<Field, FieldInfo> fieldInfos, Field field) {
    val xlsxCol = field.getAnnotation(XlsxCol.class);
    if (xlsxCol == null) {
      return;
    }

    val title = getTitle(xlsxCol);
    if (Util.isEmpty(title)) {
      return;
    }

    int i = 0;
    FieldInfo firstFieldInfo = null;
    if (!fieldInfos.isEmpty()) {
      firstFieldInfo = fieldInfos.values().iterator().next();
      i = firstFieldInfo.columnIndex() + fieldInfos.size();
    }

    FieldInfo fi = new FieldInfo();
    fieldInfos.put(field, fi);

    fi.columnIndex(i).xlsxCol(xlsxCol);

    if (styleWorkbook == null) {
      return;
    }

    String titleStyle = xlsxCol.titleStyle();
    if (Util.isNotEmpty(titleStyle)) {
      fi.titleStyle(cloneCellStyle(titleStyle));
    } else if (firstFieldInfo != null) {
      // 继承第一个注解的样式
      fi.titleStyle(firstFieldInfo.titleStyle());
    }

    String dataStyle = xlsxCol.dataStyle();
    if (Util.isNotEmpty(dataStyle)) {
      fi.dataStyle(cloneCellStyle(dataStyle));
    } else if (firstFieldInfo != null) {
      // 继承第一个注解的样式
      fi.dataStyle(firstFieldInfo.dataStyle());
    }
  }

  @SneakyThrows
  private <T> void writeCellValue(Row row, T bean, FieldInfo fi, Field field) {
    field.setAccessible(true);
    Object fieldValue = field.get(bean);
    if (fieldValue == null) {
      fieldValue = "";
    }

    Cell cell = row.createCell(fi.columnIndex());
    String titleText = fieldValue.toString();
    cell.setCellValue(titleText);

    if (fi.dataStyle() != null) {
      cell.setCellStyle(fi.dataStyle());
    }
  }

  private CellStyle cloneCellStyle(String cellReference) {
    if (styleSheet == null) {
      styleSheet = styleWorkbook.getSheetAt(0);
    }

    val cr = new CellReference(cellReference);
    val styleRow = styleSheet.getRow(cr.getRow());
    val cellStyle = styleRow.getCell(cr.getCol()).getCellStyle();
    val cloneStyle = workbook.createCellStyle();
    cloneStyle.cloneStyleFrom(cellStyle);
    return cloneStyle;
  }

  private Sheet getSheet(Class<?> beanClass) {
    return workbook.getSheetAt(0);
  }

  private String getTitle(XlsxCol xlsxCol) {
    if (Util.isNotEmpty(xlsxCol.title())) {
      return xlsxCol.title();
    }

    return xlsxCol.value();
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
      sheet = getSheet(beanClass);
    }

    Map<Field, FieldInfo> fieldInfos = createFieldFieldInfoMap(beanClass);

    for (int i = 1, ii = sheet.getLastRowNum(); i <= ii; ++i) {
      val row = sheet.getRow(i);

      T t = beanClass.getConstructor().newInstance();

      for (val entry : fieldInfos.entrySet()) {
        val field = entry.getKey();
        FieldInfo fi = entry.getValue();

        field.setAccessible(true);
        field.set(t, row.getCell(fi.columnIndex()).getStringCellValue());
      }

      beans.add(t);
    }

    return beans;
  }

  /**
   * 用密码保护工作簿。只有在xlsx格式才起作用。
   *
   * @param password 保护密码。
   * @return Xlsx
   */
  public Xlsx protectWorkbook(String password) {
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
