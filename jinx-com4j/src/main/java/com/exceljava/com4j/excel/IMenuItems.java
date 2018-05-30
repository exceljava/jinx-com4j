package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020867-0001-0000-C000-000000000046}")
public interface IMenuItems extends Com4jObject,Iterable<Com4jObject> {
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
     * <li>java.lang.Object parameter onAction is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter shortcutKey is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter before is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter restore is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter statusBar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter helpFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter helpContextID is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(caption, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param caption Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.MenuItem
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 8}, optParamIndex = {1, 2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=8)
  com.exceljava.com4j.excel.MenuItem add(
    java.lang.String caption);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter shortcutKey is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter before is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter restore is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter statusBar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter helpFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter helpContextID is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(caption, onAction, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param caption Mandatory java.lang.String parameter.
   * @param onAction Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.MenuItem
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 8}, optParamIndex = {2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=8)
  com.exceljava.com4j.excel.MenuItem add(
    java.lang.String caption,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object onAction);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter before is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter restore is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter statusBar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter helpFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter helpContextID is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(caption, onAction, shortcutKey, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param caption Mandatory java.lang.String parameter.
   * @param onAction Optional parameter. Default value is com4j.Variant.getMissing()
   * @param shortcutKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.MenuItem
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 8}, optParamIndex = {3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=8)
  com.exceljava.com4j.excel.MenuItem add(
    java.lang.String caption,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object onAction,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object shortcutKey);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter restore is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter statusBar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter helpFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter helpContextID is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(caption, onAction, shortcutKey, before, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param caption Mandatory java.lang.String parameter.
   * @param onAction Optional parameter. Default value is com4j.Variant.getMissing()
   * @param shortcutKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param before Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.MenuItem
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 8}, optParamIndex = {4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=8)
  com.exceljava.com4j.excel.MenuItem add(
    java.lang.String caption,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object onAction,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object shortcutKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object before);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter statusBar is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter helpFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter helpContextID is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(caption, onAction, shortcutKey, before, restore, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param caption Mandatory java.lang.String parameter.
   * @param onAction Optional parameter. Default value is com4j.Variant.getMissing()
   * @param shortcutKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param before Optional parameter. Default value is com4j.Variant.getMissing()
   * @param restore Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.MenuItem
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 8}, optParamIndex = {5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=8)
  com.exceljava.com4j.excel.MenuItem add(
    java.lang.String caption,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object onAction,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object shortcutKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object before,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object restore);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter helpFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter helpContextID is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(caption, onAction, shortcutKey, before, restore, statusBar, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param caption Mandatory java.lang.String parameter.
   * @param onAction Optional parameter. Default value is com4j.Variant.getMissing()
   * @param shortcutKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param before Optional parameter. Default value is com4j.Variant.getMissing()
   * @param restore Optional parameter. Default value is com4j.Variant.getMissing()
   * @param statusBar Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.MenuItem
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 8}, optParamIndex = {6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=8)
  com.exceljava.com4j.excel.MenuItem add(
    java.lang.String caption,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object onAction,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object shortcutKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object before,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object restore,
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
   * add(caption, onAction, shortcutKey, before, restore, statusBar, helpFile, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param caption Mandatory java.lang.String parameter.
   * @param onAction Optional parameter. Default value is com4j.Variant.getMissing()
   * @param shortcutKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param before Optional parameter. Default value is com4j.Variant.getMissing()
   * @param restore Optional parameter. Default value is com4j.Variant.getMissing()
   * @param statusBar Optional parameter. Default value is com4j.Variant.getMissing()
   * @param helpFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.MenuItem
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 8}, optParamIndex = {7}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=8)
  com.exceljava.com4j.excel.MenuItem add(
    java.lang.String caption,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object onAction,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object shortcutKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object before,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object restore,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object statusBar,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object helpFile);

  /**
   * @param caption Mandatory java.lang.String parameter.
   * @param onAction Optional parameter. Default value is com4j.Variant.getMissing()
   * @param shortcutKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param before Optional parameter. Default value is com4j.Variant.getMissing()
   * @param restore Optional parameter. Default value is com4j.Variant.getMissing()
   * @param statusBar Optional parameter. Default value is com4j.Variant.getMissing()
   * @param helpFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param helpContextID Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.MenuItem
   */

  @VTID(10)
  com.exceljava.com4j.excel.MenuItem add(
    java.lang.String caption,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object onAction,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object shortcutKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object before,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object restore,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object statusBar,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object helpFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object helpContextID);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter before is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter restore is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addMenu(caption, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param caption Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Menu
   */

  @VTID(11)
  @UseDefaultValues(paramIndexMapping = {0, 3}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=3)
  com.exceljava.com4j.excel.Menu addMenu(
    java.lang.String caption);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter restore is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addMenu(caption, before, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param caption Mandatory java.lang.String parameter.
   * @param before Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Menu
   */

  @VTID(11)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=3)
  com.exceljava.com4j.excel.Menu addMenu(
    java.lang.String caption,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object before);

  /**
   * @param caption Mandatory java.lang.String parameter.
   * @param before Optional parameter. Default value is com4j.Variant.getMissing()
   * @param restore Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Menu
   */

  @VTID(11)
  com.exceljava.com4j.excel.Menu addMenu(
    java.lang.String caption,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object before,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object restore);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(12)
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(13)
  @DefaultMethod
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject get_Default(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(14)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(15)
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
