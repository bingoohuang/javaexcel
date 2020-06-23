package com.github.gobars.xlsx;

import lombok.Data;
import lombok.experimental.Accessors;
import org.hibernate.validator.constraints.Length;
import org.junit.Test;

import javax.validation.constraints.NotNull;
import javax.validation.constraints.Pattern;
import java.util.ArrayList;
import java.util.List;

import static com.github.gobars.xlsx.FileType.CLASSPATH;
import static com.google.common.truth.Truth.assertThat;

public class ValidateTest {
  @Test
  public void readExcel2Beans() {
    List<Integer> errRownums = new ArrayList<>();

    Xlsx xlsx = new Xlsx().read("validate.xlsx", CLASSPATH);
    ToOption toOption =
        new ToOption()
            .errCallback(
                (XlsxValidErrCallback<XBean>)
                    (xBean, errMsg, rownum) -> {
                      errRownums.add(rownum);
                    });
    List<XBean> vBeans = xlsx.toBeans(XBean.class, toOption);
    xlsx.write("excels/test-validate.xlsx");

    assertThat(errRownums).containsExactly(6);
    assertThat(toOption.okRows()).isEqualTo(2);
    assertThat(toOption.errRows()).isEqualTo(1);
    assertThat(vBeans)
        .containsExactly(
            new XBean().area("西城").blood("135/90").gender("男"),
            new XBean().area("东城").blood("140/95").gender("女"));
  }

  @Test
  public void test() {
    VBean vb = new VBean().name("Bb1");
    assertThat(ValidateUtil.validate(vb)).isNull();

    vb = new VBean();
    assertThat(ValidateUtil.validate(vb)).isEqualTo("姓名格式错误");
  }

  @XlsxValid(writeErrorToExcel = true, removeOKRows = true)
  @Data
  @Accessors(fluent = true)
  public static class XBean implements XlsxValidatable<XBean> {
    @NotNull
    @Length(max = 3)
    @XlsxCol(title = "地区", ignoreRow = "示例-")
    private String area;

    @XlsxCol("血压")
    private String blood;

    @XlsxCol("性别")
    @Pattern(regexp = "男|女")
    private String gender;

    private String error;

    @Override
    public String validate(String error, XBean bean) {
      this.error = error;

      // 自定义其它校验，比如业务校验，状态校验、存在校验、组合校验等
      // 返回null表示校验通过，否则返回具体校验失败信息
      return error;
    }
  }

  @XlsxValid
  @Data
  @Accessors(fluent = true)
  public static class VBean implements XlsxValidatable<VBean> {
    @NotNull
    @Pattern(regexp = "[A-Z][a-z][0-9]")
    @Length(max = 3)
    @XlsxCol("姓名")
    private String name;

    private String error;

    @Override
    public String validate(String error, VBean bean) {
      this.error = error;

      // 自定义其它校验，比如业务校验，状态校验、存在校验、组合校验等
      // 返回null表示校验通过，否则返回具体校验失败信息
      return error;
    }
  }
}
