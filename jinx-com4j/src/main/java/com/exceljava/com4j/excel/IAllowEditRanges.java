package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002446A-0001-0000-C000-000000000046}")
public interface IAllowEditRanges extends Com4jObject,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(7)
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.AllowEditRange
   */

  @VTID(8)
  com.exceljava.com4j.excel.AllowEditRange getItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(title, range, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param title Mandatory java.lang.String parameter.
   * @param range Mandatory com.exceljava.com4j.excel.Range parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.AllowEditRange
   */

  @VTID(9)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=3)
  com.exceljava.com4j.excel.AllowEditRange add(
    java.lang.String title,
    com.exceljava.com4j.excel.Range range);

  /**
   * @param title Mandatory java.lang.String parameter.
   * @param range Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.AllowEditRange
   */

  @VTID(9)
  com.exceljava.com4j.excel.AllowEditRange add(
    java.lang.String title,
    com.exceljava.com4j.excel.Range range,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.AllowEditRange
   */

  @VTID(10)
  @DefaultMethod
  com.exceljava.com4j.excel.AllowEditRange get_Default(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(11)
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
