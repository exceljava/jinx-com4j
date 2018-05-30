package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface Validation extends Com4jObject {
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
     * <li>java.lang.Object parameter alertStyle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter operator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter formula1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter formula2 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(type, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlDVType parameter.
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  void add(
    com.exceljava.com4j.excel.XlDVType type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter operator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter formula1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter formula2 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(type, alertStyle, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlDVType parameter.
   * @param alertStyle Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void add(
    com.exceljava.com4j.excel.XlDVType type,
    @Optional java.lang.Object alertStyle);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter formula1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter formula2 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(type, alertStyle, operator, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlDVType parameter.
   * @param alertStyle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param operator Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void add(
    com.exceljava.com4j.excel.XlDVType type,
    @Optional java.lang.Object alertStyle,
    @Optional java.lang.Object operator);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter formula2 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(type, alertStyle, operator, formula1, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlDVType parameter.
   * @param alertStyle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param operator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula1 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void add(
    com.exceljava.com4j.excel.XlDVType type,
    @Optional java.lang.Object alertStyle,
    @Optional java.lang.Object operator,
    @Optional java.lang.Object formula1);

  /**
   * @param type Mandatory com.exceljava.com4j.excel.XlDVType parameter.
   * @param alertStyle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param operator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula2 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  void add(
    com.exceljava.com4j.excel.XlDVType type,
    @Optional java.lang.Object alertStyle,
    @Optional java.lang.Object operator,
    @Optional java.lang.Object formula1,
    @Optional java.lang.Object formula2);


  /**
   * <p>
   * Getter method for the COM property "AlertStyle"
   * </p>
   */

  @DISPID(1605)
  @PropGet
  int getAlertStyle();


  /**
   * <p>
   * Getter method for the COM property "IgnoreBlank"
   * </p>
   */

  @DISPID(1606)
  @PropGet
  boolean getIgnoreBlank();


  /**
   * <p>
   * Setter method for the COM property "IgnoreBlank"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1606)
  @PropPut
  void setIgnoreBlank(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "IMEMode"
   * </p>
   */

  @DISPID(1607)
  @PropGet
  int getIMEMode();


  /**
   * <p>
   * Setter method for the COM property "IMEMode"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(1607)
  @PropPut
  void setIMEMode(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "InCellDropdown"
   * </p>
   */

  @DISPID(1608)
  @PropGet
  boolean getInCellDropdown();


  /**
   * <p>
   * Setter method for the COM property "InCellDropdown"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1608)
  @PropPut
  void setInCellDropdown(
    boolean rhs);


  /**
   */

  @DISPID(117)
  void delete();


  /**
   * <p>
   * Getter method for the COM property "ErrorMessage"
   * </p>
   */

  @DISPID(1609)
  @PropGet
  java.lang.String getErrorMessage();


  /**
   * <p>
   * Setter method for the COM property "ErrorMessage"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1609)
  @PropPut
  void setErrorMessage(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "ErrorTitle"
   * </p>
   */

  @DISPID(1610)
  @PropGet
  java.lang.String getErrorTitle();


  /**
   * <p>
   * Setter method for the COM property "ErrorTitle"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1610)
  @PropPut
  void setErrorTitle(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "InputMessage"
   * </p>
   */

  @DISPID(1611)
  @PropGet
  java.lang.String getInputMessage();


  /**
   * <p>
   * Setter method for the COM property "InputMessage"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1611)
  @PropPut
  void setInputMessage(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "InputTitle"
   * </p>
   */

  @DISPID(1612)
  @PropGet
  java.lang.String getInputTitle();


  /**
   * <p>
   * Setter method for the COM property "InputTitle"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1612)
  @PropPut
  void setInputTitle(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Formula1"
   * </p>
   */

  @DISPID(1579)
  @PropGet
  java.lang.String getFormula1();


  /**
   * <p>
   * Getter method for the COM property "Formula2"
   * </p>
   */

  @DISPID(1580)
  @PropGet
  java.lang.String getFormula2();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter alertStyle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter operator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter formula1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter formula2 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * modify(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(1581)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  void modify();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter alertStyle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter operator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter formula1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter formula2 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * modify(type, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1581)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  void modify(
    @Optional java.lang.Object type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter operator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter formula1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter formula2 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * modify(type, alertStyle, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param alertStyle Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1581)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void modify(
    @Optional java.lang.Object type,
    @Optional java.lang.Object alertStyle);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter formula1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter formula2 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * modify(type, alertStyle, operator, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param alertStyle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param operator Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1581)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void modify(
    @Optional java.lang.Object type,
    @Optional java.lang.Object alertStyle,
    @Optional java.lang.Object operator);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter formula2 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * modify(type, alertStyle, operator, formula1, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param alertStyle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param operator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula1 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1581)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void modify(
    @Optional java.lang.Object type,
    @Optional java.lang.Object alertStyle,
    @Optional java.lang.Object operator,
    @Optional java.lang.Object formula1);

  /**
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param alertStyle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param operator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula2 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1581)
  void modify(
    @Optional java.lang.Object type,
    @Optional java.lang.Object alertStyle,
    @Optional java.lang.Object operator,
    @Optional java.lang.Object formula1,
    @Optional java.lang.Object formula2);


  /**
   * <p>
   * Getter method for the COM property "Operator"
   * </p>
   */

  @DISPID(797)
  @PropGet
  int getOperator();


  /**
   * <p>
   * Getter method for the COM property "ShowError"
   * </p>
   */

  @DISPID(1613)
  @PropGet
  boolean getShowError();


  /**
   * <p>
   * Setter method for the COM property "ShowError"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1613)
  @PropPut
  void setShowError(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowInput"
   * </p>
   */

  @DISPID(1614)
  @PropGet
  boolean getShowInput();


  /**
   * <p>
   * Setter method for the COM property "ShowInput"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1614)
  @PropPut
  void setShowInput(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   */

  @DISPID(108)
  @PropGet
  int getType();


  /**
   * <p>
   * Getter method for the COM property "Value"
   * </p>
   */

  @DISPID(6)
  @PropGet
  boolean getValue();


  // Properties:
}
