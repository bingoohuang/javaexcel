package com.github.gobars.xlsx;

import lombok.Data;
import lombok.experimental.Accessors;
import org.junit.Test;

import java.util.ArrayList;
import java.util.List;

import static com.google.common.truth.Truth.assertThat;

public class XlsxTest {
  @Test
  public void rows() {
    List<RowBean> rows = new ArrayList<>();
    rows.add(new RowBean().name("黄进兵"));

    String excel = "excels/test-rowsbean.xlsx";
    new Xlsx().fromBeans(rows).write(excel);
    List<RowBean> read = new Xlsx().read(excel).toBeans(RowBean.class);

    assertThat(read).isEqualTo(rows);
  }

  @Test
  public void rows2() {
    List<RowBean2> rows = new ArrayList<>();
    rows.add(new RowBean2().name("黄进兵"));

    String excel = "excels/test-rowsbean2.xlsx";

    new Xlsx().styleXlsx("rowsbean2-style.xlsx", FileType.CLASSPATH).fromBeans(rows).write(excel);

    List<RowBean2> read = new Xlsx().read(excel).toBeans(RowBean2.class);

    assertThat(read).isEqualTo(rows);
  }

  @Test
  public void rows3() {
    List<RowBean3> rows = new ArrayList<>();
    rows.add(new RowBean3().name("黄进兵").city("海淀区"));
    rows.add(new RowBean3().name("兵进黄").city("西城区"));

    String excel = "excels/test-rowsbean3.xlsx";

    new Xlsx().styleXlsx("rowsbean2-style.xlsx", FileType.CLASSPATH).fromBeans(rows).write(excel);

    List<RowBean3> read = new Xlsx().read(excel).toBeans(RowBean3.class);

    assertThat(read).isEqualTo(rows);
  }

  @Data
  @Accessors(fluent = true)
  public static class RowBean {
    @XlsxCol("姓名")
    private String name;
  }

  @Data
  @Accessors(fluent = true)
  public static class RowBean2 {
    @XlsxCol(title = "姓名", titleStyle = "A1", dataStyle = "A2")
    private String name;
  }

  @Data
  @Accessors(fluent = true)
  public static class RowBean3 {
    @XlsxCol(title = "姓名", titleStyle = "A1", dataStyle = "A2")
    private String name;

    @XlsxCol(title = "城市")
    private String city;
  }
}
