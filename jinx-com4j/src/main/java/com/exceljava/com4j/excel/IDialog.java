package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002087A-0001-0000-C000-000000000046}")
public interface IDialog extends Com4jObject {
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
     * <li>java.lang.Object parameter arg1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg3 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg4 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg5 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg6 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg7 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg8 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg9 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg10 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg11 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg12 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg13 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg14 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {30}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg3 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg4 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg5 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg6 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg7 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg8 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg9 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg10 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg11 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg12 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg13 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg14 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 30}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg3 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg4 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg5 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg6 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg7 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg8 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg9 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg10 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg11 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg12 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg13 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg14 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, arg2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 30}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg4 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg5 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg6 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg7 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg8 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg9 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg10 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg11 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg12 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg13 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg14 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, arg2, arg3, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 30}, optParamIndex = {3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg5 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg6 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg7 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg8 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg9 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg10 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg11 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg12 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg13 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg14 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, arg2, arg3, arg4, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 30}, optParamIndex = {4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg6 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg7 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg8 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg9 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg10 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg11 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg12 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg13 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg14 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, arg2, arg3, arg4, arg5, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 30}, optParamIndex = {5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg7 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg8 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg9 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg10 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg11 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg12 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg13 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg14 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, arg2, arg3, arg4, arg5, arg6, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 30}, optParamIndex = {6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg8 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg9 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg10 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg11 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg12 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg13 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg14 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, arg2, arg3, arg4, arg5, arg6, arg7, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 30}, optParamIndex = {7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg9 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg10 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg11 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg12 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg13 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg14 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 30}, optParamIndex = {8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg10 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg11 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg12 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg13 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg14 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 30}, optParamIndex = {9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg11 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg12 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg13 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg14 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 30}, optParamIndex = {10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg12 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg13 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg14 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 30}, optParamIndex = {11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg13 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg14 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 30}, optParamIndex = {12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg14 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 30}, optParamIndex = {13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg15 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 30}, optParamIndex = {14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg16 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 30}, optParamIndex = {15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg17 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 30}, optParamIndex = {16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg18 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg17 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 30}, optParamIndex = {17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg17);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg19 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg17 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg18 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 30}, optParamIndex = {18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg17,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg18);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg20 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg17 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg18 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg19 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 30}, optParamIndex = {19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg17,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg18,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg19);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg21 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg17 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg18 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg19 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg20 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 30}, optParamIndex = {20, 21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg17,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg18,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg19,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg20);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg22 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg17 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg18 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg19 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg20 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg21 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 30}, optParamIndex = {21, 22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg17,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg18,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg19,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg20,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg21);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg23 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg17 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg18 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg19 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg20 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg21 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg22 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 30}, optParamIndex = {22, 23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg17,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg18,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg19,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg20,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg21,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg22);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg24 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg17 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg18 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg19 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg20 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg21 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg22 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg23 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 30}, optParamIndex = {23, 24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg17,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg18,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg19,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg20,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg21,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg22,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg23);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg25 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg17 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg18 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg19 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg20 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg21 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg22 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg23 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg24 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 30}, optParamIndex = {24, 25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg17,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg18,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg19,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg20,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg21,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg22,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg23,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg24);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg26 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg17 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg18 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg19 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg20 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg21 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg22 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg23 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg24 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg25 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 30}, optParamIndex = {25, 26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg17,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg18,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg19,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg20,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg21,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg22,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg23,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg24,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg25);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg27 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg17 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg18 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg19 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg20 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg21 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg22 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg23 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg24 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg25 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg26 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 30}, optParamIndex = {26, 27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg17,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg18,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg19,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg20,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg21,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg22,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg23,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg24,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg25,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg26);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg28 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg17 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg18 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg19 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg20 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg21 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg22 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg23 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg24 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg25 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg26 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg27 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 30}, optParamIndex = {27, 28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg17,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg18,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg19,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg20,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg21,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg22,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg23,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg24,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg25,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg26,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg27);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg29 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg17 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg18 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg19 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg20 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg21 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg22 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg23 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg24 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg25 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg26 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg27 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg28 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 30}, optParamIndex = {28, 29}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg17,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg18,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg19,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg20,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg21,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg22,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg23,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg24,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg25,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg26,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg27,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg28);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter arg30 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg17 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg18 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg19 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg20 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg21 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg22 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg23 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg24 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg25 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg26 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg27 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg28 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg29 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 30}, optParamIndex = {29}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=30)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg17,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg18,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg19,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg20,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg21,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg22,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg23,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg24,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg25,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg26,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg27,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg28,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg29);

  /**
   * @param arg1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg5 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg6 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg7 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg8 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg9 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg10 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg11 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg12 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg13 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg14 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg15 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg16 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg17 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg18 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg19 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg20 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg21 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg22 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg23 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg24 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg25 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg26 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg27 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg28 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg29 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param arg30 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  boolean show(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg4,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg5,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg6,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg7,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg8,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg9,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg10,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg11,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg12,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg13,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg14,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg15,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg16,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg17,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg18,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg19,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg20,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg21,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg22,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg23,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg24,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg25,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg26,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg27,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg28,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg29,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object arg30);


  // Properties:
}
