package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244D0-0001-0000-C000-000000000046}")
public interface IQuickAnalysis extends Com4jObject {
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
     * <li>com.exceljava.com4j.excel.XlQuickAnalysisMode parameter xlQuickAnalysisMode is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(0);
   * </code>
   * </p>
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {com.exceljava.com4j.excel.XlQuickAnalysisMode.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"0"})
  void show();

  /**
   * @param xlQuickAnalysisMode Optional parameter. Default value is 0
   */

  @VTID(10)
  void show(
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlQuickAnalysisMode xlQuickAnalysisMode);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlQuickAnalysisMode parameter xlQuickAnalysisMode is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * hide(0);
   * </code>
   * </p>
   */

  @VTID(11)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {com.exceljava.com4j.excel.XlQuickAnalysisMode.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"0"})
  void hide();

  /**
   * @param xlQuickAnalysisMode Optional parameter. Default value is 0
   */

  @VTID(11)
  void hide(
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlQuickAnalysisMode xlQuickAnalysisMode);


  // Properties:
}
