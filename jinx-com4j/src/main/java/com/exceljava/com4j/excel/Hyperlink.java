package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface Hyperlink extends Com4jObject {
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
   * Getter method for the COM property "Name"
   * </p>
   */

  @DISPID(110)
  @PropGet
  java.lang.String getName();


  /**
   * <p>
   * Getter method for the COM property "Range"
   * </p>
   */

  @DISPID(197)
  @PropGet
  com.exceljava.com4j.excel.Range getRange();


  /**
   * <p>
   * Getter method for the COM property "Shape"
   * </p>
   */

  @DISPID(1582)
  @PropGet
  com.exceljava.com4j.excel.Shape getShape();


  /**
   * <p>
   * Getter method for the COM property "SubAddress"
   * </p>
   */

  @DISPID(1471)
  @PropGet
  java.lang.String getSubAddress();


  /**
   * <p>
   * Setter method for the COM property "SubAddress"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1471)
  @PropPut
  void setSubAddress(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Address"
   * </p>
   */

  @DISPID(236)
  @PropGet
  java.lang.String getAddress();


  /**
   * <p>
   * Setter method for the COM property "Address"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(236)
  @PropPut
  void setAddress(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   */

  @DISPID(108)
  @PropGet
  int getType();


  /**
   */

  @DISPID(1476)
  void addToFavorites();


  /**
   */

  @DISPID(117)
  void delete();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter newWindow is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addHistory is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter extraInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter method is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter headerInfo is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * follow(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(1616)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  void follow();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter addHistory is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter extraInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter method is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter headerInfo is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * follow(newWindow, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param newWindow Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1616)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  void follow(
    @Optional java.lang.Object newWindow);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter extraInfo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter method is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter headerInfo is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * follow(newWindow, addHistory, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param newWindow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addHistory Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1616)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void follow(
    @Optional java.lang.Object newWindow,
    @Optional java.lang.Object addHistory);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter method is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter headerInfo is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * follow(newWindow, addHistory, extraInfo, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param newWindow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addHistory Optional parameter. Default value is com4j.Variant.getMissing()
   * @param extraInfo Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1616)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void follow(
    @Optional java.lang.Object newWindow,
    @Optional java.lang.Object addHistory,
    @Optional java.lang.Object extraInfo);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter headerInfo is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * follow(newWindow, addHistory, extraInfo, method, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param newWindow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addHistory Optional parameter. Default value is com4j.Variant.getMissing()
   * @param extraInfo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param method Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1616)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void follow(
    @Optional java.lang.Object newWindow,
    @Optional java.lang.Object addHistory,
    @Optional java.lang.Object extraInfo,
    @Optional java.lang.Object method);

  /**
   * @param newWindow Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addHistory Optional parameter. Default value is com4j.Variant.getMissing()
   * @param extraInfo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param method Optional parameter. Default value is com4j.Variant.getMissing()
   * @param headerInfo Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1616)
  void follow(
    @Optional java.lang.Object newWindow,
    @Optional java.lang.Object addHistory,
    @Optional java.lang.Object extraInfo,
    @Optional java.lang.Object method,
    @Optional java.lang.Object headerInfo);


  /**
   * <p>
   * Getter method for the COM property "EmailSubject"
   * </p>
   */

  @DISPID(1883)
  @PropGet
  java.lang.String getEmailSubject();


  /**
   * <p>
   * Setter method for the COM property "EmailSubject"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1883)
  @PropPut
  void setEmailSubject(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "ScreenTip"
   * </p>
   */

  @DISPID(1881)
  @PropGet
  java.lang.String getScreenTip();


  /**
   * <p>
   * Setter method for the COM property "ScreenTip"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1881)
  @PropPut
  void setScreenTip(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "TextToDisplay"
   * </p>
   */

  @DISPID(1882)
  @PropGet
  java.lang.String getTextToDisplay();


  /**
   * <p>
   * Setter method for the COM property "TextToDisplay"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1882)
  @PropPut
  void setTextToDisplay(
    java.lang.String rhs);


  /**
   * @param filename Mandatory java.lang.String parameter.
   * @param editNow Mandatory boolean parameter.
   * @param overwrite Mandatory boolean parameter.
   */

  @DISPID(1884)
  void createNewDocument(
    java.lang.String filename,
    boolean editNow,
    boolean overwrite);


  // Properties:
}
