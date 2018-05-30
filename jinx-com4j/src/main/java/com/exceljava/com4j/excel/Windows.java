package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface Windows extends Com4jObject,Iterable<Com4jObject> {
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
     * <li>com.exceljava.com4j.excel.XlArrangeStyle parameter arrangeStyle is set to 1</li><li>java.lang.Object parameter activeWorkbook is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter syncHorizontal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter syncVertical is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * arrange(1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(638)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3}, javaType = {com.exceljava.com4j.excel.XlArrangeStyle.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"1", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object arrange();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter activeWorkbook is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter syncHorizontal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter syncVertical is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * arrange(arrangeStyle, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arrangeStyle Optional parameter. Default value is 1
   */

  @DISPID(638)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object arrange(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlArrangeStyle arrangeStyle);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter syncHorizontal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter syncVertical is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * arrange(arrangeStyle, activeWorkbook, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arrangeStyle Optional parameter. Default value is 1
   * @param activeWorkbook Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(638)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object arrange(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlArrangeStyle arrangeStyle,
    @Optional java.lang.Object activeWorkbook);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter syncVertical is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * arrange(arrangeStyle, activeWorkbook, syncHorizontal, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arrangeStyle Optional parameter. Default value is 1
   * @param activeWorkbook Optional parameter. Default value is com4j.Variant.getMissing()
   * @param syncHorizontal Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(638)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object arrange(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlArrangeStyle arrangeStyle,
    @Optional java.lang.Object activeWorkbook,
    @Optional java.lang.Object syncHorizontal);

  /**
   * @param arrangeStyle Optional parameter. Default value is 1
   * @param activeWorkbook Optional parameter. Default value is com4j.Variant.getMissing()
   * @param syncHorizontal Optional parameter. Default value is com4j.Variant.getMissing()
   * @param syncVertical Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(638)
  java.lang.Object arrange(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlArrangeStyle arrangeStyle,
    @Optional java.lang.Object activeWorkbook,
    @Optional java.lang.Object syncHorizontal,
    @Optional java.lang.Object syncVertical);


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
  com.exceljava.com4j.excel.Window getItem(
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
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   */

  @DISPID(0)
  @PropGet
  @DefaultMethod
  com.exceljava.com4j.excel.Window get_Default(
    java.lang.Object index);


  /**
   * @param windowName Mandatory java.lang.Object parameter.
   */

  @DISPID(2246)
  boolean compareSideBySideWith(
    java.lang.Object windowName);


  /**
   */

  @DISPID(2248)
  boolean breakSideBySide();


  /**
   * <p>
   * Getter method for the COM property "SyncScrollingSideBySide"
   * </p>
   */

  @DISPID(2249)
  @PropGet
  boolean getSyncScrollingSideBySide();


  /**
   * <p>
   * Setter method for the COM property "SyncScrollingSideBySide"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2249)
  @PropPut
  void setSyncScrollingSideBySide(
    boolean rhs);


  /**
   */

  @DISPID(2250)
  void resetPositionsSideBySide();


  // Properties:
}
