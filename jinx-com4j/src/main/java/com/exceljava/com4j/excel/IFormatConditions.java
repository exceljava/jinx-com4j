package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024424-0001-0000-C000-000000000046}")
public interface IFormatConditions extends Com4jObject,Iterable<Com4jObject> {
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
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(10)
  int getCount();


  /**
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(11)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject item(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


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
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(12)
  @UseDefaultValues(paramIndexMapping = {0, 8}, optParamIndex = {1, 2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.Dispatch,index=8)
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
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(12)
  @UseDefaultValues(paramIndexMapping = {0, 1, 8}, optParamIndex = {2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.Dispatch,index=8)
  com4j.Com4jObject add(
    com.exceljava.com4j.excel.XlFormatConditionType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object operator);

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
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(12)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 8}, optParamIndex = {3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.Dispatch,index=8)
  com4j.Com4jObject add(
    com.exceljava.com4j.excel.XlFormatConditionType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object operator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object formula1);

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
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(12)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 8}, optParamIndex = {4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.Dispatch,index=8)
  com4j.Com4jObject add(
    com.exceljava.com4j.excel.XlFormatConditionType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object operator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object formula1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object formula2);

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
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(12)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 8}, optParamIndex = {5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.Dispatch,index=8)
  com4j.Com4jObject add(
    com.exceljava.com4j.excel.XlFormatConditionType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object operator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object formula1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object formula2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object string);

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
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(12)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 8}, optParamIndex = {6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.Dispatch,index=8)
  com4j.Com4jObject add(
    com.exceljava.com4j.excel.XlFormatConditionType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object operator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object formula1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object formula2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object string,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textOperator);

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
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(12)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 8}, optParamIndex = {7}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.Dispatch,index=8)
  com4j.Com4jObject add(
    com.exceljava.com4j.excel.XlFormatConditionType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object operator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object formula1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object formula2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object string,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textOperator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dateOperator);

  /**
   * @param type Mandatory com.exceljava.com4j.excel.XlFormatConditionType parameter.
   * @param operator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param formula2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param string Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textOperator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dateOperator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param scopeType Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(12)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject add(
    com.exceljava.com4j.excel.XlFormatConditionType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object operator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object formula1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object formula2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object string,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textOperator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dateOperator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object scopeType);


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
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(14)
  java.util.Iterator<Com4jObject> iterator();

  /**
   */

  @VTID(15)
  void delete();


  /**
   * @param colorScaleType Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(16)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject addColorScale(
    int colorScaleType);


  /**
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(17)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject addDatabar();


  /**
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(18)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject addIconSetCondition();


  /**
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(19)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject addTop10();


  /**
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(20)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject addAboveAverage();


  /**
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(21)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject addUniqueValues();


  // Properties:
}
