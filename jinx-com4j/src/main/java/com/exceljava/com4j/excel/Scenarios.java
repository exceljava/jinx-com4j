package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface Scenarios extends Com4jObject,Iterable<Com4jObject> {
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
     * <li>java.lang.Object parameter values is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comment is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter locked is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hidden is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, changingCells, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param changingCells Mandatory java.lang.Object parameter.
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Scenario add(
    java.lang.String name,
    java.lang.Object changingCells);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter comment is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter locked is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hidden is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, changingCells, values, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param changingCells Mandatory java.lang.Object parameter.
   * @param values Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Scenario add(
    java.lang.String name,
    java.lang.Object changingCells,
    @Optional java.lang.Object values);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter locked is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hidden is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, changingCells, values, comment, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param changingCells Mandatory java.lang.Object parameter.
   * @param values Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comment Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Scenario add(
    java.lang.String name,
    java.lang.Object changingCells,
    @Optional java.lang.Object values,
    @Optional java.lang.Object comment);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter hidden is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, changingCells, values, comment, locked, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param changingCells Mandatory java.lang.Object parameter.
   * @param values Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comment Optional parameter. Default value is com4j.Variant.getMissing()
   * @param locked Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Scenario add(
    java.lang.String name,
    java.lang.Object changingCells,
    @Optional java.lang.Object values,
    @Optional java.lang.Object comment,
    @Optional java.lang.Object locked);

  /**
   * @param name Mandatory java.lang.String parameter.
   * @param changingCells Mandatory java.lang.Object parameter.
   * @param values Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comment Optional parameter. Default value is com4j.Variant.getMissing()
   * @param locked Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hidden Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  com.exceljava.com4j.excel.Scenario add(
    java.lang.String name,
    java.lang.Object changingCells,
    @Optional java.lang.Object values,
    @Optional java.lang.Object comment,
    @Optional java.lang.Object locked,
    @Optional java.lang.Object hidden);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   */

  @DISPID(118)
  @PropGet
  int getCount();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlSummaryReportType parameter reportType is set to 1</li><li>java.lang.Object parameter resultCells is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createSummary(1, com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(913)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {com.exceljava.com4j.excel.XlSummaryReportType.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"1", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object createSummary();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter resultCells is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createSummary(reportType, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param reportType Optional parameter. Default value is 1
   */

  @DISPID(913)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object createSummary(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSummaryReportType reportType);

  /**
   * @param reportType Optional parameter. Default value is 1
   * @param resultCells Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(913)
  java.lang.Object createSummary(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSummaryReportType reportType,
    @Optional java.lang.Object resultCells);


  /**
   * @param index Mandatory java.lang.Object parameter.
   */

  @DISPID(170)
  com.exceljava.com4j.excel.Scenario item(
    java.lang.Object index);


  /**
   * @param source Mandatory java.lang.Object parameter.
   */

  @DISPID(564)
  java.lang.Object merge(
    java.lang.Object source);


  /**
   */

  @DISPID(-4)
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
