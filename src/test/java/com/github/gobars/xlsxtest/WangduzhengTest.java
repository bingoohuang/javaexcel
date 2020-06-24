package com.github.gobars.xlsxtest;

import com.github.gobars.xlsx.Xlsx;
import com.github.gobars.xlsx.XlsxCol;
import com.github.gobars.xlsx.XlsxOptionFrom;
import com.github.gobars.xlsx.XlsxUtil;
import lombok.Data;
import lombok.experimental.Accessors;
import org.junit.Test;

import java.util.List;

import static com.github.gobars.xlsx.XlsxFileType.CLASSPATH;

public class WangduzhengTest {
  @Test
  public void test() {
    List<Title> bs =
        XlsxUtil.listOf(
            new Title().title("登录名").sample1("阿三").sample2("阿四"),
            new Title().title("证件号码"),
            new Title().title("手机号"),
            new Title().title("身份标签(xxx)"),
            new Title().title("所属机构"),
            new Title().title("一个属性"),
            new Title().title("两个属性"),
            new Title().title("N个属性").sample1("aa").sample2("bb"));

    new Xlsx()
        .read("identity_import_src1.xlsx", CLASSPATH)
        .fromBeans(bs, new XlsxOptionFrom().horizontal(true))
        .write("excels/用户输入模板.xlsx");

    new Xlsx()
        .read("identity_import_src1.xls", CLASSPATH)
        .fromBeans(bs, new XlsxOptionFrom().horizontal(true))
        .write("excels/用户输入模板.xls");
  }

  @Data
  @Accessors(fluent = true)
  public class Title {
    @XlsxCol("标题")
    private String title;

    @XlsxCol("示例1")
    private String sample1;

    @XlsxCol("示例2")
    private String sample2;
  }
}
