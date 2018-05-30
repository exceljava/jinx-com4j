package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface QuickAnalysis extends Com4jObject {
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
     * <li>com.exceljava.com4j.excel.XlQuickAnalysisMode parameter xlQuickAnalysisMode is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(0);
   * </code>
   * </p>
   */

  @DISPID(496)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {com.exceljava.com4j.excel.XlQuickAnalysisMode.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"0"})
  void show();

  /**
   * @param xlQuickAnalysisMode Optional parameter. Default value is 0
   */

  @DISPID(496)
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

  @DISPID(813)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {com.exceljava.com4j.excel.XlQuickAnalysisMode.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"0"})
  void hide();

  /**
   * @param xlQuickAnalysisMode Optional parameter. Default value is 0
   */

  @DISPID(813)
  void hide(
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlQuickAnalysisMode xlQuickAnalysisMode);


  // Properties:
}
