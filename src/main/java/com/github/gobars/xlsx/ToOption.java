package com.github.gobars.xlsx;

import lombok.Data;
import lombok.experimental.Accessors;

@Data
@Accessors(fluent = true)
public class ToOption {
  int okRows;
  int errRows;

  XlsxValidErrCallback errCallback;
}
