package com.github.gobars.xlsx;

import lombok.experimental.UtilityClass;

@UtilityClass
public class Util {
  public boolean isEmpty(String s) {
    return s == null || s.length() == 0;
  }

  public boolean isNotEmpty(String s) {
    return s != null && s.length() > 0;
  }
}
