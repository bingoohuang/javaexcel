package com.github.gobars.xlsxtest;

import com.github.gobars.xlsx.XlsxUtil;
import org.junit.Test;

import static com.google.common.truth.Truth.assertThat;

public class XlsxUtilTest {
  @Test
  public void testFmtDouble() {
    assertThat(XlsxUtil.fmt(1.0)).isEqualTo("1");
    assertThat(XlsxUtil.fmt(1.01)).isEqualTo("1.01");
    assertThat(XlsxUtil.fmt(999999999)).isEqualTo("999999999");
  }
}
