package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface ModelMeasures extends Com4jObject,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   */

  @DISPID(148)
  @PropGet
  com.exceljava.com4j.excel._Application getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   */

  @DISPID(149)
  @PropGet
  com.exceljava.com4j.excel.XlCreator getCreator();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   */

  @DISPID(150)
  @PropGet
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   */

  @DISPID(118)
  @PropGet
  int getCount();


  /**
   * @param index Mandatory java.lang.Object parameter.
   */

  @DISPID(170)
  com.exceljava.com4j.excel.ModelMeasure item(
    java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   */

  @DISPID(0)
  @PropGet
  @DefaultMethod
  com.exceljava.com4j.excel.ModelMeasure get_Default(
    java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4)
  @PropGet
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter description is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(measureName, associatedTable, formula, formatInformation, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param measureName Mandatory java.lang.String parameter.
   * @param associatedTable Mandatory com.exceljava.com4j.excel.ModelTable parameter.
   * @param formula Mandatory java.lang.String parameter.
   * @param formatInformation Mandatory java.lang.Object parameter.
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.ModelMeasure add(
    java.lang.String measureName,
    com.exceljava.com4j.excel.ModelTable associatedTable,
    java.lang.String formula,
    java.lang.Object formatInformation);

  /**
   * @param measureName Mandatory java.lang.String parameter.
   * @param associatedTable Mandatory com.exceljava.com4j.excel.ModelTable parameter.
   * @param formula Mandatory java.lang.String parameter.
   * @param formatInformation Mandatory java.lang.Object parameter.
   * @param description Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  com.exceljava.com4j.excel.ModelMeasure add(
    java.lang.String measureName,
    com.exceljava.com4j.excel.ModelTable associatedTable,
    java.lang.String formula,
    java.lang.Object formatInformation,
    @Optional java.lang.Object description);


  // Properties:
}
