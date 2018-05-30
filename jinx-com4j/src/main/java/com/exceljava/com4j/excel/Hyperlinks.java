package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface Hyperlinks extends Com4jObject,Iterable<Com4jObject> {
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
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter subAddress is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter screenTip is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textToDisplay is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(anchor, address, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param anchor Mandatory com4j.Com4jObject parameter.
   * @param address Mandatory java.lang.String parameter.
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com4j.Com4jObject add(
    com4j.Com4jObject anchor,
    java.lang.String address);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter screenTip is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textToDisplay is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(anchor, address, subAddress, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param anchor Mandatory com4j.Com4jObject parameter.
   * @param address Mandatory java.lang.String parameter.
   * @param subAddress Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  com4j.Com4jObject add(
    com4j.Com4jObject anchor,
    java.lang.String address,
    @Optional java.lang.Object subAddress);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter textToDisplay is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(anchor, address, subAddress, screenTip, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param anchor Mandatory com4j.Com4jObject parameter.
   * @param address Mandatory java.lang.String parameter.
   * @param subAddress Optional parameter. Default value is com4j.Variant.getMissing()
   * @param screenTip Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com4j.Com4jObject add(
    com4j.Com4jObject anchor,
    java.lang.String address,
    @Optional java.lang.Object subAddress,
    @Optional java.lang.Object screenTip);

  /**
   * @param anchor Mandatory com4j.Com4jObject parameter.
   * @param address Mandatory java.lang.String parameter.
   * @param subAddress Optional parameter. Default value is com4j.Variant.getMissing()
   * @param screenTip Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textToDisplay Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  com4j.Com4jObject add(
    com4j.Com4jObject anchor,
    java.lang.String address,
    @Optional java.lang.Object subAddress,
    @Optional java.lang.Object screenTip,
    @Optional java.lang.Object textToDisplay);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   */

  @DISPID(118)
  @PropGet
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   */

  @DISPID(170)
  @PropGet
  com.exceljava.com4j.excel.Hyperlink getItem(
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
  com.exceljava.com4j.excel.Hyperlink get_Default(
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
   */

  @DISPID(117)
  void delete();


  // Properties:
}
