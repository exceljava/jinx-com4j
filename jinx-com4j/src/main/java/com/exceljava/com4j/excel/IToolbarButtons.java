package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002085F-0001-0000-C000-000000000046}")
public interface IToolbarButtons extends Com4jObject,Iterable<Com4jObject> {
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
     * <li>java.lang.Object parameter button is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter before is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter onAction is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pushed is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter enabled is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter statusBar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter helpFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter helpContextID is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ToolbarButton
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {8}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=8)
  com.exceljava.com4j.excel.ToolbarButton add();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter before is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter onAction is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pushed is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter enabled is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter statusBar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter helpFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter helpContextID is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(button, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param button Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.ToolbarButton
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 8}, optParamIndex = {1, 2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=8)
  com.exceljava.com4j.excel.ToolbarButton add(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object button);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter onAction is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pushed is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter enabled is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter statusBar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter helpFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter helpContextID is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(button, before, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param button Optional parameter. Default value is com4j.Variant.getMissing()
   * @param before Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.ToolbarButton
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 8}, optParamIndex = {2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=8)
  com.exceljava.com4j.excel.ToolbarButton add(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object button,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object before);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter pushed is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter enabled is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter statusBar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter helpFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter helpContextID is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(button, before, onAction, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param button Optional parameter. Default value is com4j.Variant.getMissing()
   * @param before Optional parameter. Default value is com4j.Variant.getMissing()
   * @param onAction Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.ToolbarButton
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 8}, optParamIndex = {3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=8)
  com.exceljava.com4j.excel.ToolbarButton add(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object button,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object before,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object onAction);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter enabled is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter statusBar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter helpFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter helpContextID is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(button, before, onAction, pushed, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param button Optional parameter. Default value is com4j.Variant.getMissing()
   * @param before Optional parameter. Default value is com4j.Variant.getMissing()
   * @param onAction Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pushed Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.ToolbarButton
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 8}, optParamIndex = {4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=8)
  com.exceljava.com4j.excel.ToolbarButton add(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object button,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object before,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object onAction,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pushed);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter statusBar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter helpFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter helpContextID is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(button, before, onAction, pushed, enabled, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param button Optional parameter. Default value is com4j.Variant.getMissing()
   * @param before Optional parameter. Default value is com4j.Variant.getMissing()
   * @param onAction Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pushed Optional parameter. Default value is com4j.Variant.getMissing()
   * @param enabled Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.ToolbarButton
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 8}, optParamIndex = {5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=8)
  com.exceljava.com4j.excel.ToolbarButton add(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object button,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object before,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object onAction,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pushed,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object enabled);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter helpFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter helpContextID is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(button, before, onAction, pushed, enabled, statusBar, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param button Optional parameter. Default value is com4j.Variant.getMissing()
   * @param before Optional parameter. Default value is com4j.Variant.getMissing()
   * @param onAction Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pushed Optional parameter. Default value is com4j.Variant.getMissing()
   * @param enabled Optional parameter. Default value is com4j.Variant.getMissing()
   * @param statusBar Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.ToolbarButton
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 8}, optParamIndex = {6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=8)
  com.exceljava.com4j.excel.ToolbarButton add(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object button,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object before,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object onAction,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pushed,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object enabled,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object statusBar);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter helpContextID is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(button, before, onAction, pushed, enabled, statusBar, helpFile, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param button Optional parameter. Default value is com4j.Variant.getMissing()
   * @param before Optional parameter. Default value is com4j.Variant.getMissing()
   * @param onAction Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pushed Optional parameter. Default value is com4j.Variant.getMissing()
   * @param enabled Optional parameter. Default value is com4j.Variant.getMissing()
   * @param statusBar Optional parameter. Default value is com4j.Variant.getMissing()
   * @param helpFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.ToolbarButton
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 8}, optParamIndex = {7}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=8)
  com.exceljava.com4j.excel.ToolbarButton add(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object button,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object before,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object onAction,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pushed,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object enabled,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object statusBar,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object helpFile);

  /**
   * @param button Optional parameter. Default value is com4j.Variant.getMissing()
   * @param before Optional parameter. Default value is com4j.Variant.getMissing()
   * @param onAction Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pushed Optional parameter. Default value is com4j.Variant.getMissing()
   * @param enabled Optional parameter. Default value is com4j.Variant.getMissing()
   * @param statusBar Optional parameter. Default value is com4j.Variant.getMissing()
   * @param helpFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param helpContextID Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.ToolbarButton
   */

  @VTID(10)
  com.exceljava.com4j.excel.ToolbarButton add(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object button,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object before,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object onAction,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pushed,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object enabled,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object statusBar,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object helpFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object helpContextID);


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
   * @param index Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.ToolbarButton
   */

  @VTID(12)
  com.exceljava.com4j.excel.ToolbarButton getItem(
    int index);


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
   * @param index Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.ToolbarButton
   */

  @VTID(14)
  @DefaultMethod
  com.exceljava.com4j.excel.ToolbarButton get_Default(
    int index);


  // Properties:
}
