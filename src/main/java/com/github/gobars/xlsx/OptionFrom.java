package com.github.gobars.xlsx;

import lombok.Data;
import lombok.experimental.Accessors;

@Data
@Accessors(fluent = true)
public class OptionFrom {
  /** 是否水平写（从左到右写). */
  private boolean horizontal;

  private String titleStyle;
  private String dataStyle;
}
