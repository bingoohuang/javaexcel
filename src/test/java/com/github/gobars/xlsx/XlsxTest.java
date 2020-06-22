package com.github.gobars.xlsx;

import lombok.Data;
import lombok.experimental.Accessors;
import org.junit.Test;

import java.util.ArrayList;
import java.util.List;

import static com.google.common.truth.Truth.assertThat;

public class XlsxTest {
  @Data
  @Accessors(fluent = true)
  public static class RowBean {
    @XlsxCol("姓名")
    private String name;
  }

  @Test
  public void rows() {
    List<RowBean> rows = new ArrayList<>();
    rows.add(new RowBean().name("黄进兵"));

    new Xlsx().writeBeans(rows).write("rowsbean.xlsx");

    List<RowBean> read = new Xlsx().read("rowsbean.xlsx").readBeans(RowBean.class);

    assertThat(read).isEqualTo(rows);
  }
}
