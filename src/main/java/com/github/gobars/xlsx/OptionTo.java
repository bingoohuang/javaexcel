package com.github.gobars.xlsx;

import lombok.Data;
import lombok.experimental.Accessors;

@Data
@Accessors(fluent = true)
public class OptionTo {
  private int okRows;
  private int errRows;

  private XlsxValidErrCallback errCallback;

  private boolean writeErrorToExcel;
  private boolean removeOkRows;

  private XlsxIgnoreCallback ignoreCallback;

  private XlsxValidatable validatable;
}
