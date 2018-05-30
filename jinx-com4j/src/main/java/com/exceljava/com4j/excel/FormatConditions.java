package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface FormatConditions extends Com4jObject,Iterable<Com4jObject> {
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
   * Getter method for the COM property "Count"
   * </p>
   */

  @DISPID(118)
  @PropGet
  int getCount();


  /**
   * @param index Mandatory java.lang.Object parameter.
   */

  @DISPID(170)
  com4j.Com4jObject item(
    java.lang.Object index);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter operator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter formula1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter formula2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter string is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textOperator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter dateOperator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter scopeType is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(type, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFormatConditionType parameter.
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com4j.Com4jObject add(
    com.exceljava.com4j.excel.XlFormatConditionType type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter formula1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter formula2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter string is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textOperator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter dateOperator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter scopeType is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(type, operator, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFormatConditionType parameter.
   * @param operator Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com4j.Com4jObject add(
    com.exceljava.com4j.excel.XlFormatConditionType type,
    @Optional java.lang.Object operator);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter formula2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter string is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textOperator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter dateOperator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter scopeType is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(type, operator, formula1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFormatConditionType parameter.
   * @param operator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula1 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com4j.Com4jObject add(
    com.exceljava.com4j.excel.XlFormatConditionType type,
    @Optional java.lang.Object operator,
    @Optional java.lang.Object formula1);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter string is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textOperator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter dateOperator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter scopeType is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(type, operator, formula1, formula2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFormatConditionType parameter.
   * @param operator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula2 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com4j.Com4jObject add(
    com.exceljava.com4j.excel.XlFormatConditionType type,
    @Optional java.lang.Object operator,
    @Optional java.lang.Object formula1,
    @Optional java.lang.Object formula2);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter textOperator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter dateOperator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter scopeType is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(type, operator, formula1, formula2, string, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFormatConditionType parameter.
   * @param operator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param string Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com4j.Com4jObject add(
    com.exceljava.com4j.excel.XlFormatConditionType type,
    @Optional java.lang.Object operator,
    @Optional java.lang.Object formula1,
    @Optional java.lang.Object formula2,
    @Optional java.lang.Object string);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter dateOperator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter scopeType is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(type, operator, formula1, formula2, string, textOperator, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFormatConditionType parameter.
   * @param operator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param string Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textOperator Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  com4j.Com4jObject add(
    com.exceljava.com4j.excel.XlFormatConditionType type,
    @Optional java.lang.Object operator,
    @Optional java.lang.Object formula1,
    @Optional java.lang.Object formula2,
    @Optional java.lang.Object string,
    @Optional java.lang.Object textOperator);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter scopeType is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(type, operator, formula1, formula2, string, textOperator, dateOperator, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFormatConditionType parameter.
   * @param operator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param string Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textOperator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dateOperator Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com4j.Com4jObject add(
    com.exceljava.com4j.excel.XlFormatConditionType type,
    @Optional java.lang.Object operator,
    @Optional java.lang.Object formula1,
    @Optional java.lang.Object formula2,
    @Optional java.lang.Object string,
    @Optional java.lang.Object textOperator,
    @Optional java.lang.Object dateOperator);

  /**
   * @param type Mandatory com.exceljava.com4j.excel.XlFormatConditionType parameter.
   * @param operator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param string Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textOperator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dateOperator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param scopeType Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  com4j.Com4jObject add(
    com.exceljava.com4j.excel.XlFormatConditionType type,
    @Optional java.lang.Object operator,
    @Optional java.lang.Object formula1,
    @Optional java.lang.Object formula2,
    @Optional java.lang.Object string,
    @Optional java.lang.Object textOperator,
    @Optional java.lang.Object dateOperator,
    @Optional java.lang.Object scopeType);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   */

  @DISPID(0)
  @PropGet
  @DefaultMethod
  com4j.Com4jObject get_Default(
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


  /**
   * @param colorScaleType Mandatory int parameter.
   */

  @DISPID(2616)
  com4j.Com4jObject addColorScale(
    int colorScaleType);


  /**
   */

  @DISPID(2618)
  com4j.Com4jObject addDatabar();


  /**
   */

  @DISPID(2619)
  com4j.Com4jObject addIconSetCondition();


  /**
   */

  @DISPID(2620)
  com4j.Com4jObject addTop10();


  /**
   */

  @DISPID(2621)
  com4j.Com4jObject addAboveAverage();


  /**
   */

  @DISPID(2622)
  com4j.Com4jObject addUniqueValues();


  // Properties:
}
