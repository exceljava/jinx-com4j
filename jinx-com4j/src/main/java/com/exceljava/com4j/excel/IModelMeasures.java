package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244EE-0001-0000-C000-000000000046}")
public interface IModelMeasures extends Com4jObject,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel._Application
   */

  @VTID(7)
  com.exceljava.com4j.excel._Application getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlCreator
   */

  @VTID(8)
  com.exceljava.com4j.excel.XlCreator getCreator();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(9)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(10)
  int getCount();


  /**
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelMeasure
   */

  @VTID(11)
  com.exceljava.com4j.excel.ModelMeasure item(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelMeasure
   */

  @VTID(12)
  @DefaultMethod
  com.exceljava.com4j.excel.ModelMeasure get_Default(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(13)
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
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelMeasure
   */

  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 5}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.excel.ModelMeasure add(
    java.lang.String measureName,
    com.exceljava.com4j.excel.ModelTable associatedTable,
    java.lang.String formula,
    @MarshalAs(NativeType.VARIANT) java.lang.Object formatInformation);

  /**
   * @param measureName Mandatory java.lang.String parameter.
   * @param associatedTable Mandatory com.exceljava.com4j.excel.ModelTable parameter.
   * @param formula Mandatory java.lang.String parameter.
   * @param formatInformation Mandatory java.lang.Object parameter.
   * @param description Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelMeasure
   */

  @VTID(14)
  com.exceljava.com4j.excel.ModelMeasure add(
    java.lang.String measureName,
    com.exceljava.com4j.excel.ModelTable associatedTable,
    java.lang.String formula,
    @MarshalAs(NativeType.VARIANT) java.lang.Object formatInformation,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object description);


  // Properties:
}
