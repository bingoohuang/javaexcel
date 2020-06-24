package com.github.gobars.xlsx;

import lombok.Cleanup;
import lombok.SneakyThrows;
import lombok.experimental.UtilityClass;
import lombok.val;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileInputStream;
import java.io.InputStream;

@UtilityClass
public class XlsxReader {
  @SneakyThrows
  public Workbook read(String fileName, XlsxFileType fileType) {
    switch (fileType) {
      case NORMAL:
        @Cleanup val fis = new FileInputStream(fileName);
        return read(fis);
      case CLASSPATH:
        @Cleanup val cis = Xlsx.class.getClassLoader().getResourceAsStream(fileName);
        if (cis == null) {
          throw new XlsxException("unable to find excel file in classpath " + fileName);
        }

        return read(cis);
      default:
        throw new XlsxException("unsupported fileType " + fileType);
    }
  }

  @SneakyThrows
  public Workbook read(InputStream is) {
    return WorkbookFactory.create(is);
  }
}
