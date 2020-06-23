package com.github.gobars.xlsx;

import java.lang.annotation.*;

/**
 * Variant of JSR-303's {@link javax.validation.Valid}, supporting the specification of validation
 * groups. Designed for convenient use with Spring's JSR-303 support but not JSR-303 specific.
 *
 * <p>Can also be used with method level validation, indicating that a specific class is supposed to
 * be validated at the method level (acting as a pointcut for the corresponding validation
 * interceptor), but also optionally specifying the validation groups for method-level validation in
 * the annotated class. Applying this annotation at the method level allows for overriding the
 * validation groups for a specific method but does not serve as a pointcut; a class-level
 * annotation is nevertheless necessary to trigger method validation for a specific bean to begin
 * with. Can also be used as a meta-annotation on a custom stereotype annotation or a custom
 * group-specific validated annotation.
 *
 * @author Juergen Hoeller
 * @see javax.validation.Validator#validate(Object, Class[])
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.TYPE, ElementType.METHOD, ElementType.PARAMETER})
public @interface XlsxValid {
  /**
   * 是否将错误写回Excel.
   *
   * @return true 是
   */
  boolean writeErrorToExcel() default false;

  /**
   * 是否删除校验通过的数据行.
   *
   * @return true 是
   */
  boolean removeOKRows() default false;
}
