package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020896-0001-0000-C000-000000000046}")
public interface IScenarios extends Com4jObject,Iterable<Com4jObject> {
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
     * <li>java.lang.Object parameter values is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter comment is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter locked is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hidden is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, changingCells, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param changingCells Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Scenario
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 6}, optParamIndex = {2, 3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=6)
  com.exceljava.com4j.excel.Scenario add(
    java.lang.String name,
    @MarshalAs(NativeType.VARIANT) java.lang.Object changingCells);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.Scenario
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 6}, optParamIndex = {3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=6)
  com.exceljava.com4j.excel.Scenario add(
    java.lang.String name,
    @MarshalAs(NativeType.VARIANT) java.lang.Object changingCells,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object values);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.Scenario
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 6}, optParamIndex = {4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=6)
  com.exceljava.com4j.excel.Scenario add(
    java.lang.String name,
    @MarshalAs(NativeType.VARIANT) java.lang.Object changingCells,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object values,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comment);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.Scenario
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 6}, optParamIndex = {5}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=6)
  com.exceljava.com4j.excel.Scenario add(
    java.lang.String name,
    @MarshalAs(NativeType.VARIANT) java.lang.Object changingCells,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object values,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comment,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object locked);

  /**
   * @param name Mandatory java.lang.String parameter.
   * @param changingCells Mandatory java.lang.Object parameter.
   * @param values Optional parameter. Default value is com4j.Variant.getMissing()
   * @param comment Optional parameter. Default value is com4j.Variant.getMissing()
   * @param locked Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hidden Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Scenario
   */

  @VTID(10)
  com.exceljava.com4j.excel.Scenario add(
    java.lang.String name,
    @MarshalAs(NativeType.VARIANT) java.lang.Object changingCells,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object values,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object comment,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object locked,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hidden);


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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(12)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {com.exceljava.com4j.excel.XlSummaryReportType.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"1", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(12)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
  java.lang.Object createSummary(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSummaryReportType reportType);

  /**
   * @param reportType Optional parameter. Default value is 1
   * @param resultCells Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(12)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object createSummary(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlSummaryReportType reportType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object resultCells);


  /**
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Scenario
   */

  @VTID(13)
  com.exceljava.com4j.excel.Scenario item(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * @param source Mandatory java.lang.Object parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(14)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object merge(
    @MarshalAs(NativeType.VARIANT) java.lang.Object source);


  /**
   */

  @VTID(15)
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
