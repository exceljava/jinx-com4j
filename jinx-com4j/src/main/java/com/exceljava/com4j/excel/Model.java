package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface Model extends Com4jObject {
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
   * Getter method for the COM property "ModelTables"
   * </p>
   */

  @DISPID(3106)
  @PropGet
  com.exceljava.com4j.excel.ModelTables getModelTables();


  /**
   * <p>
   * Getter method for the COM property "ModelRelationships"
   * </p>
   */

  @DISPID(3126)
  @PropGet
  com.exceljava.com4j.excel.ModelRelationships getModelRelationships();


  /**
   */

  @DISPID(1417)
  void refresh();


  /**
   * @param connectionToDataSource Mandatory com.exceljava.com4j.excel.WorkbookConnection parameter.
   */

  @DISPID(3127)
  com.exceljava.com4j.excel.WorkbookConnection addConnection(
    com.exceljava.com4j.excel.WorkbookConnection connectionToDataSource);


  /**
   * @param modelTable Mandatory java.lang.Object parameter.
   */

  @DISPID(3129)
  com.exceljava.com4j.excel.WorkbookConnection createModelWorkbookConnection(
    java.lang.Object modelTable);


  /**
   * <p>
   * Getter method for the COM property "DataModelConnection"
   * </p>
   */

  @DISPID(3131)
  @PropGet
  com.exceljava.com4j.excel.WorkbookConnection getDataModelConnection();


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   */

  @DISPID(110)
  @PropGet
  java.lang.String getName();


  /**
   */

  @DISPID(3132)
  void initialize();


  /**
   * <p>
   * Getter method for the COM property "ModelMeasures"
   * </p>
   */

  @DISPID(3210)
  @PropGet
  com.exceljava.com4j.excel.ModelMeasures getModelMeasures();


  /**
   * <p>
   * Getter method for the COM property "ModelFormatGeneral"
   * </p>
   */

  @DISPID(3211)
  @PropGet
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
   */

  @DISPID(3212)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.ModelFormatDate getModelFormatDate();

  /**
   * <p>
   * Getter method for the COM property "ModelFormatDate"
   * </p>
   * @param formatString Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3212)
  @PropGet
  com.exceljava.com4j.excel.ModelFormatDate getModelFormatDate(
    @Optional java.lang.Object formatString);


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
   */

  @DISPID(3214)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
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
   */

  @DISPID(3214)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.ModelFormatDecimalNumber getModelFormatDecimalNumber(
    @Optional java.lang.Object useThousandSeparator);

  /**
   * <p>
   * Getter method for the COM property "ModelFormatDecimalNumber"
   * </p>
   * @param useThousandSeparator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param decimalPlaces Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3214)
  @PropGet
  com.exceljava.com4j.excel.ModelFormatDecimalNumber getModelFormatDecimalNumber(
    @Optional java.lang.Object useThousandSeparator,
    @Optional java.lang.Object decimalPlaces);


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
   */

  @DISPID(3216)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.ModelFormatWholeNumber getModelFormatWholeNumber();

  /**
   * <p>
   * Getter method for the COM property "ModelFormatWholeNumber"
   * </p>
   * @param useThousandSeparator Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3216)
  @PropGet
  com.exceljava.com4j.excel.ModelFormatWholeNumber getModelFormatWholeNumber(
    @Optional java.lang.Object useThousandSeparator);


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
   */

  @DISPID(3217)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
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
   */

  @DISPID(3217)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.ModelFormatPercentageNumber getModelFormatPercentageNumber(
    @Optional java.lang.Object useThousandSeparator);

  /**
   * <p>
   * Getter method for the COM property "ModelFormatPercentageNumber"
   * </p>
   * @param useThousandSeparator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param decimalPlaces Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3217)
  @PropGet
  com.exceljava.com4j.excel.ModelFormatPercentageNumber getModelFormatPercentageNumber(
    @Optional java.lang.Object useThousandSeparator,
    @Optional java.lang.Object decimalPlaces);


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
   */

  @DISPID(3218)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.ModelFormatScientificNumber getModelFormatScientificNumber();

  /**
   * <p>
   * Getter method for the COM property "ModelFormatScientificNumber"
   * </p>
   * @param decimalPlaces Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3218)
  @PropGet
  com.exceljava.com4j.excel.ModelFormatScientificNumber getModelFormatScientificNumber(
    @Optional java.lang.Object decimalPlaces);


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
   */

  @DISPID(3219)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
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
   */

  @DISPID(3219)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.ModelFormatCurrency getModelFormatCurrency(
    @Optional java.lang.Object symbol);

  /**
   * <p>
   * Getter method for the COM property "ModelFormatCurrency"
   * </p>
   * @param symbol Optional parameter. Default value is com4j.Variant.getMissing()
   * @param decimalPlaces Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3219)
  @PropGet
  com.exceljava.com4j.excel.ModelFormatCurrency getModelFormatCurrency(
    @Optional java.lang.Object symbol,
    @Optional java.lang.Object decimalPlaces);


  /**
   * <p>
   * Getter method for the COM property "ModelFormatBoolean"
   * </p>
   */

  @DISPID(3221)
  @PropGet
  com.exceljava.com4j.excel.ModelFormatBoolean getModelFormatBoolean();


  // Properties:
}
