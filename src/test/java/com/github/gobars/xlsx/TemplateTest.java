package com.github.gobars.xlsx;

import lombok.Data;
import lombok.experimental.Accessors;
import org.junit.Test;

import java.util.ArrayList;
import java.util.List;

import static com.google.common.truth.Truth.assertThat;

public class TemplateTest {
  @Test
  public void test() {
    List<XlsxTest.RowBean3> rows = new ArrayList<>();
    rows.add(new XlsxTest.RowBean3().name("黄进兵").city("海淀区"));
    rows.add(new XlsxTest.RowBean3().name("兵进黄").city("西城区"));

    String excel = "excels/test-template.xlsx";

    new Xlsx().templateXlsx("template.xlsx", FileType.CLASSPATH).fromBeans(rows).write(excel);

    List<XlsxTest.RowBean3> read = new Xlsx().read(excel).toBeans(XlsxTest.RowBean3.class);

    assertThat(read).isEqualTo(rows);
  }

  @Data
  @Accessors(fluent = true)
  public static class RowBean3 {
    @XlsxCol(title = "姓名")
    private String name;

    @XlsxCol(title = "城市")
    private String city;
  }
}
