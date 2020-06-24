package com.github.gobars.xlsx;

import lombok.Data;
import lombok.experimental.Accessors;

@Data
@Accessors(fluent = true)
public class XlsxOptionTo {
  private int okRows;
  private int errRows;

  private XlsxValidErrable errable;

  private boolean writeErrorToExcel;
  private boolean removeOkRows;

  private XlsxIgnoreable ignoreable;

  private XlsxValidatable validatable;
}
