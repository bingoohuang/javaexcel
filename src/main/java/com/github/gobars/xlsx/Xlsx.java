package com.github.gobars.xlsx;

import lombok.Cleanup;
import lombok.SneakyThrows;
import lombok.val;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.List;

public class Xlsx implements Closeable {
  private Workbook workbook;
  private Sheet sheet;

  /**
   * 写入列表.
   *
   * @param rows bean列表
   * @return Xlsx
   */
  @SneakyThrows
  public <T> Xlsx writeBeans(List<T> rows) {
    if (rows.size() == 0) {
      return this;
    }

    val beanClass = rows.get(0).getClass();

    if (workbook == null) {
      workbook = new XSSFWorkbook();
      sheet = workbook.createSheet();
    } else {
      sheet = getSheet(beanClass);
    }

    Row row = sheet.createRow(0);
    int i = 0;

    for (val field : beanClass.getDeclaredFields()) {
      val title = getTitle(field.getAnnotation(XlsxCol.class));
      if (Util.isEmpty(title)) {
        continue;
      }

      Cell cell = row.createCell(i);
      cell.setCellValue(title);
      i++;
    }

    for (val bean : rows) {
      row = sheet.createRow(sheet.getLastRowNum() + 1);

      int j = 0;
      for (val field : beanClass.getDeclaredFields()) {
        val title = getTitle(field.getAnnotation(XlsxCol.class));
        if (Util.isEmpty(title)) {
          continue;
        }

        field.setAccessible(true);
        Object fieldValue = field.get(bean);
        if (fieldValue == null) {
          fieldValue = "";
        }

        Cell cell = row.createCell(j);
        cell.setCellValue(fieldValue.toString());
        j++;
      }
    }

    return this;
  }

  private Sheet getSheet(Class<?> beanClass) {
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
   * @param beanClass JavaBean类型
   * @param <T> JavaBean类型
   * @return JavaBean列表
   */
  @SneakyThrows
  public <T> List<T> readBeans(Class<T> beanClass) {
    ArrayList<T> beans = new ArrayList<>(10);

    if (sheet == null) {
      sheet = getSheet(beanClass);
    }

    for (int i = 1, ii = sheet.getLastRowNum(); i <= ii; ++i) {
      val row = sheet.getRow(i);

      T t = beanClass.getConstructor().newInstance();

      for (val field : beanClass.getDeclaredFields()) {
        val title = getTitle(field.getAnnotation(XlsxCol.class));
        if (Util.isEmpty(title)) {
          continue;
        }

        field.setAccessible(true);
        field.set(t, row.getCell(0).getStringCellValue());
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
    switch (fileType) {
      case NORMAL:
        @Cleanup val fis = new FileInputStream(fileName);
        return read(fis);
      case CLASSPATH:
        @Cleanup val cis = Xlsx.class.getClassLoader().getResourceAsStream(fileName);
        if (cis != null) {
          throw new XlsxException("unable to find excel file in classpath " + fileName);
        }

        return read(cis);
      default:
        throw new XlsxException("unsupported fileType " + fileType);
    }
  }

  @SneakyThrows
  public Xlsx read(String fileName) {
    return read(fileName, FileType.NORMAL);
  }

  @SneakyThrows
  public Xlsx read(InputStream is) {
    this.workbook = WorkbookFactory.create(is);
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
