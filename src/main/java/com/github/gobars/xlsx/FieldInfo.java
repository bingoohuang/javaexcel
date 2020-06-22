package com.github.gobars.xlsx;

import lombok.Data;
import lombok.experimental.Accessors;
import org.apache.poi.ss.usermodel.CellStyle;

@Data
@Accessors(fluent = true)
public class FieldInfo {
  CellStyle titleStyle;
  CellStyle dataStyle;
  int columnIndex;
  XlsxCol xlsxCol;
}
