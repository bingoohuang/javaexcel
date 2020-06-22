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
    List<RowBean3> beans = new ArrayList<>();
    beans.add(new RowBean3().name("黄进兵").city("海淀区"));
    beans.add(new RowBean3().name("兵进黄").city("西城区"));

    String excel = "excels/test-template.xlsx";

    new Xlsx().templateXlsx("template.xlsx", FileType.CLASSPATH).fromBeans(beans).write(excel);

    List<RowBean3> read = new Xlsx().read(excel).toBeans(RowBean3.class);

    assertThat(read).isEqualTo(beans);
  }

  @Test
  public void horizontal() {
    List<TitleBean> beans = new ArrayList<>();
    beans.add(new TitleBean().title("地区").sample("示例-海淀区").data1("西城").data2("东城").data3("南城"));
    beans.add(
        new TitleBean()
            .title("血压")
            .sample("示例-140/90")
            .data1("135/90")
            .data2("140/95")
            .data3("133/85"));
    beans.add(new TitleBean().title("性别").sample("示例-女").data1("男").data2("女").data3("未知"));
    beans.add(new TitleBean().title("学校").sample("示例-蓝翔").data1("东大").data2("西大").data3("北大"));

    String excel = "excels/test-horizontal.xlsx";
    new Xlsx()
        .templateXlsx("template-horizontal.xlsx", FileType.CLASSPATH)
        .fromBeans(beans, new FromOption().horizontal(true))
        .write(excel);

    List<HorizontalBean> read = new Xlsx().read(excel).toBeans(HorizontalBean.class);
    assertThat(read)
        .containsExactly(
            new HorizontalBean()
                .area("示例-海淀区")
                .bloodPressure("示例-140/90")
                .gender("示例-女")
                .school("示例-蓝翔"),
            new HorizontalBean().area("西城").bloodPressure("135/90").gender("男").school("东大"),
            new HorizontalBean().area("东城").bloodPressure("140/95").gender("女").school("西大"),
            new HorizontalBean().area("南城").bloodPressure("133/85").gender("未知").school("北大"));
  }

  @Data
  @Accessors(fluent = true)
  public static class HorizontalBean {
    @XlsxCol(title = "地区")
    private String area;

    @XlsxCol(title = "血压")
    private String bloodPressure;

    @XlsxCol(title = "性别")
    private String gender;

    @XlsxCol(title = "学校")
    private String school;
  }

  @Data
  @Accessors(fluent = true)
  public static class TitleBean {
    @XlsxCol(title = "数据1")
    private String data1;

    @XlsxCol(title = "数据2")
    private String data2;

    @XlsxCol(title = "数据3")
    private String data3;

    @XlsxCol(title = "标题")
    private String title;

    @XlsxCol(title = "示例")
    private String sample;
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
