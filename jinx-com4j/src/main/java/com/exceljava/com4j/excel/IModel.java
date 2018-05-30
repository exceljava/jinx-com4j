package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244DB-0001-0000-C000-000000000046}")
public interface IModel extends Com4jObject {
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
   * Getter method for the COM property "ModelTables"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelTables
   */

  @VTID(10)
  com.exceljava.com4j.excel.ModelTables getModelTables();


  /**
   * <p>
   * Getter method for the COM property "ModelRelationships"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelRelationships
   */

  @VTID(11)
  com.exceljava.com4j.excel.ModelRelationships getModelRelationships();


  /**
   */

  @VTID(12)
  void refresh();


  /**
   * @param connectionToDataSource Mandatory com.exceljava.com4j.excel.WorkbookConnection parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.WorkbookConnection
   */

  @VTID(13)
  com.exceljava.com4j.excel.WorkbookConnection addConnection(
    com.exceljava.com4j.excel.WorkbookConnection connectionToDataSource);


  /**
   * @param modelTable Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.WorkbookConnection
   */

  @VTID(14)
  com.exceljava.com4j.excel.WorkbookConnection createModelWorkbookConnection(
    @MarshalAs(NativeType.VARIANT) java.lang.Object modelTable);


  /**
   * <p>
   * Getter method for the COM property "DataModelConnection"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.WorkbookConnection
   */

  @VTID(15)
  com.exceljava.com4j.excel.WorkbookConnection getDataModelConnection();


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(16)
  java.lang.String getName();


  /**
   */

  @VTID(17)
  void initialize();


  /**
   * <p>
   * Getter method for the COM property "ModelMeasures"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelMeasures
   */

  @VTID(18)
  com.exceljava.com4j.excel.ModelMeasures getModelMeasures();


  /**
   * <p>
   * Getter method for the COM property "ModelFormatGeneral"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelFormatGeneral
   */

  @VTID(19)
  com.exceljava.com4j.excel.ModelFormatGeneral getModelFormatGeneral();


  /**
   * <p>
   * Getter method for the COM property "ModelFormatDate"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter formatString is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getModelFormatDate(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelFormatDate
   */

  @VTID(20)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.ModelFormatDate getModelFormatDate();

  /**
   * <p>
   * Getter method for the COM property "ModelFormatDate"
   * </p>
   * @param formatString Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelFormatDate
   */

  @VTID(20)
  com.exceljava.com4j.excel.ModelFormatDate getModelFormatDate(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object formatString);


  /**
   * <p>
   * Getter method for the COM property "ModelFormatDecimalNumber"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter useThousandSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalPlaces is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getModelFormatDecimalNumber(com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelFormatDecimalNumber
   */

  @VTID(21)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=2)
  com.exceljava.com4j.excel.ModelFormatDecimalNumber getModelFormatDecimalNumber();

  /**
   * <p>
   * Getter method for the COM property "ModelFormatDecimalNumber"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter decimalPlaces is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getModelFormatDecimalNumber(useThousandSeparator, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param useThousandSeparator Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelFormatDecimalNumber
   */

  @VTID(21)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=2)
  com.exceljava.com4j.excel.ModelFormatDecimalNumber getModelFormatDecimalNumber(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object useThousandSeparator);

  /**
   * <p>
   * Getter method for the COM property "ModelFormatDecimalNumber"
   * </p>
   * @param useThousandSeparator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param decimalPlaces Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelFormatDecimalNumber
   */

  @VTID(21)
  com.exceljava.com4j.excel.ModelFormatDecimalNumber getModelFormatDecimalNumber(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object useThousandSeparator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object decimalPlaces);


  /**
   * <p>
   * Getter method for the COM property "ModelFormatWholeNumber"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter useThousandSeparator is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getModelFormatWholeNumber(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelFormatWholeNumber
   */

  @VTID(22)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.ModelFormatWholeNumber getModelFormatWholeNumber();

  /**
   * <p>
   * Getter method for the COM property "ModelFormatWholeNumber"
   * </p>
   * @param useThousandSeparator Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelFormatWholeNumber
   */

  @VTID(22)
  com.exceljava.com4j.excel.ModelFormatWholeNumber getModelFormatWholeNumber(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object useThousandSeparator);


  /**
   * <p>
   * Getter method for the COM property "ModelFormatPercentageNumber"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter useThousandSeparator is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalPlaces is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getModelFormatPercentageNumber(com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelFormatPercentageNumber
   */

  @VTID(23)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=2)
  com.exceljava.com4j.excel.ModelFormatPercentageNumber getModelFormatPercentageNumber();

  /**
   * <p>
   * Getter method for the COM property "ModelFormatPercentageNumber"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter decimalPlaces is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getModelFormatPercentageNumber(useThousandSeparator, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param useThousandSeparator Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelFormatPercentageNumber
   */

  @VTID(23)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=2)
  com.exceljava.com4j.excel.ModelFormatPercentageNumber getModelFormatPercentageNumber(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object useThousandSeparator);

  /**
   * <p>
   * Getter method for the COM property "ModelFormatPercentageNumber"
   * </p>
   * @param useThousandSeparator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param decimalPlaces Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelFormatPercentageNumber
   */

  @VTID(23)
  com.exceljava.com4j.excel.ModelFormatPercentageNumber getModelFormatPercentageNumber(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object useThousandSeparator,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object decimalPlaces);


  /**
   * <p>
   * Getter method for the COM property "ModelFormatScientificNumber"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter decimalPlaces is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getModelFormatScientificNumber(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelFormatScientificNumber
   */

  @VTID(24)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.ModelFormatScientificNumber getModelFormatScientificNumber();

  /**
   * <p>
   * Getter method for the COM property "ModelFormatScientificNumber"
   * </p>
   * @param decimalPlaces Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelFormatScientificNumber
   */

  @VTID(24)
  com.exceljava.com4j.excel.ModelFormatScientificNumber getModelFormatScientificNumber(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object decimalPlaces);


  /**
   * <p>
   * Getter method for the COM property "ModelFormatCurrency"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter symbol is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter decimalPlaces is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getModelFormatCurrency(com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelFormatCurrency
   */

  @VTID(25)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=2)
  com.exceljava.com4j.excel.ModelFormatCurrency getModelFormatCurrency();

  /**
   * <p>
   * Getter method for the COM property "ModelFormatCurrency"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter decimalPlaces is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getModelFormatCurrency(symbol, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param symbol Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelFormatCurrency
   */

  @VTID(25)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=2)
  com.exceljava.com4j.excel.ModelFormatCurrency getModelFormatCurrency(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object symbol);

  /**
   * <p>
   * Getter method for the COM property "ModelFormatCurrency"
   * </p>
   * @param symbol Optional parameter. Default value is com4j.Variant.getMissing()
   * @param decimalPlaces Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelFormatCurrency
   */

  @VTID(25)
  com.exceljava.com4j.excel.ModelFormatCurrency getModelFormatCurrency(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object symbol,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object decimalPlaces);


  /**
   * <p>
   * Getter method for the COM property "ModelFormatBoolean"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelFormatBoolean
   */

  @VTID(26)
  com.exceljava.com4j.excel.ModelFormatBoolean getModelFormatBoolean();


  // Properties:
}
