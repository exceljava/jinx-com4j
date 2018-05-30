package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020892-0001-0000-C000-000000000046}")
public interface IWindows extends Com4jObject,Iterable<Com4jObject> {
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {4}, optParamIndex = {0, 1, 2, 3}, javaType = {com.exceljava.com4j.excel.XlArrangeStyle.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"1", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 4}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 4}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object arrange(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlArrangeStyle arrangeStyle,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activeWorkbook);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 4}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object arrange(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlArrangeStyle arrangeStyle,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activeWorkbook,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object syncHorizontal);

  /**
   * @param arrangeStyle Optional parameter. Default value is 1
   * @param activeWorkbook Optional parameter. Default value is com4j.Variant.getMissing()
   * @param syncHorizontal Optional parameter. Default value is com4j.Variant.getMissing()
   * @param syncVertical Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(10)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object arrange(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlArrangeStyle arrangeStyle,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activeWorkbook,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object syncHorizontal,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object syncVertical);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(11)
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Window
   */

  @VTID(12)
  com.exceljava.com4j.excel.Window getItem(
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
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Window
   */

  @VTID(14)
  @DefaultMethod
  com.exceljava.com4j.excel.Window get_Default(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * @param windowName Mandatory java.lang.Object parameter.
   * @return  Returns a value of type boolean
   */

  @VTID(15)
  boolean compareSideBySideWith(
    @MarshalAs(NativeType.VARIANT) java.lang.Object windowName);


  /**
   * @return  Returns a value of type boolean
   */

  @VTID(16)
  boolean breakSideBySide();


  /**
   * <p>
   * Getter method for the COM property "SyncScrollingSideBySide"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(17)
  boolean getSyncScrollingSideBySide();


  /**
   * <p>
   * Setter method for the COM property "SyncScrollingSideBySide"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(18)
  void setSyncScrollingSideBySide(
    boolean rhs);


  /**
   */

  @VTID(19)
  void resetPositionsSideBySide();


  // Properties:
}
