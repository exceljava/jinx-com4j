package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface SeriesCollection extends Com4jObject,Iterable<Com4jObject> {
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
     * <li>com.exceljava.com4j.excel.XlRowCol parameter rowcol is set to -4105</li><li>java.lang.Object parameter seriesLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter categoryLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter replace is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(source, -4105, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param source Mandatory java.lang.Object parameter.
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4}, javaType = {com.exceljava.com4j.excel.XlRowCol.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"-4105", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Series add(
    java.lang.Object source);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter seriesLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter categoryLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter replace is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(source, rowcol, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param source Mandatory java.lang.Object parameter.
   * @param rowcol Optional parameter. Default value is -4105
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Series add(
    java.lang.Object source,
    @Optional @DefaultValue("-4105") com.exceljava.com4j.excel.XlRowCol rowcol);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter categoryLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter replace is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(source, rowcol, seriesLabels, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param source Mandatory java.lang.Object parameter.
   * @param rowcol Optional parameter. Default value is -4105
   * @param seriesLabels Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Series add(
    java.lang.Object source,
    @Optional @DefaultValue("-4105") com.exceljava.com4j.excel.XlRowCol rowcol,
    @Optional java.lang.Object seriesLabels);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter replace is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(source, rowcol, seriesLabels, categoryLabels, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param source Mandatory java.lang.Object parameter.
   * @param rowcol Optional parameter. Default value is -4105
   * @param seriesLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param categoryLabels Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Series add(
    java.lang.Object source,
    @Optional @DefaultValue("-4105") com.exceljava.com4j.excel.XlRowCol rowcol,
    @Optional java.lang.Object seriesLabels,
    @Optional java.lang.Object categoryLabels);

  /**
   * @param source Mandatory java.lang.Object parameter.
   * @param rowcol Optional parameter. Default value is -4105
   * @param seriesLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param categoryLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param replace Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  com.exceljava.com4j.excel.Series add(
    java.lang.Object source,
    @Optional @DefaultValue("-4105") com.exceljava.com4j.excel.XlRowCol rowcol,
    @Optional java.lang.Object seriesLabels,
    @Optional java.lang.Object categoryLabels,
    @Optional java.lang.Object replace);


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
     * <li>java.lang.Object parameter rowcol is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter categoryLabels is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * extend(source, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param source Mandatory java.lang.Object parameter.
   */

  @DISPID(227)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object extend(
    java.lang.Object source);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter categoryLabels is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * extend(source, rowcol, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param source Mandatory java.lang.Object parameter.
   * @param rowcol Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(227)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object extend(
    java.lang.Object source,
    @Optional java.lang.Object rowcol);

  /**
   * @param source Mandatory java.lang.Object parameter.
   * @param rowcol Optional parameter. Default value is com4j.Variant.getMissing()
   * @param categoryLabels Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(227)
  java.lang.Object extend(
    java.lang.Object source,
    @Optional java.lang.Object rowcol,
    @Optional java.lang.Object categoryLabels);


  /**
   * @param index Mandatory java.lang.Object parameter.
   */

  @DISPID(170)
  com.exceljava.com4j.excel.Series item(
    java.lang.Object index);


  /**
   */

  @DISPID(-4)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlRowCol parameter rowcol is set to -4105</li><li>java.lang.Object parameter seriesLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter categoryLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter replace is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter newSeries is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * paste(-4105, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(211)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4}, javaType = {com.exceljava.com4j.excel.XlRowCol.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"-4105", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object paste();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter seriesLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter categoryLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter replace is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter newSeries is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * paste(rowcol, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param rowcol Optional parameter. Default value is -4105
   */

  @DISPID(211)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object paste(
    @Optional @DefaultValue("-4105") com.exceljava.com4j.excel.XlRowCol rowcol);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter categoryLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter replace is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter newSeries is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * paste(rowcol, seriesLabels, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param rowcol Optional parameter. Default value is -4105
   * @param seriesLabels Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(211)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object paste(
    @Optional @DefaultValue("-4105") com.exceljava.com4j.excel.XlRowCol rowcol,
    @Optional java.lang.Object seriesLabels);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter replace is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter newSeries is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * paste(rowcol, seriesLabels, categoryLabels, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param rowcol Optional parameter. Default value is -4105
   * @param seriesLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param categoryLabels Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(211)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.Object paste(
    @Optional @DefaultValue("-4105") com.exceljava.com4j.excel.XlRowCol rowcol,
    @Optional java.lang.Object seriesLabels,
    @Optional java.lang.Object categoryLabels);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter newSeries is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * paste(rowcol, seriesLabels, categoryLabels, replace, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param rowcol Optional parameter. Default value is -4105
   * @param seriesLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param categoryLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param replace Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(211)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object paste(
    @Optional @DefaultValue("-4105") com.exceljava.com4j.excel.XlRowCol rowcol,
    @Optional java.lang.Object seriesLabels,
    @Optional java.lang.Object categoryLabels,
    @Optional java.lang.Object replace);

  /**
   * @param rowcol Optional parameter. Default value is -4105
   * @param seriesLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param categoryLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param replace Optional parameter. Default value is com4j.Variant.getMissing()
   * @param newSeries Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(211)
  java.lang.Object paste(
    @Optional @DefaultValue("-4105") com.exceljava.com4j.excel.XlRowCol rowcol,
    @Optional java.lang.Object seriesLabels,
    @Optional java.lang.Object categoryLabels,
    @Optional java.lang.Object replace,
    @Optional java.lang.Object newSeries);


  /**
   */

  @DISPID(1117)
  com.exceljava.com4j.excel.Series newSeries();


  /**
   * @param index Mandatory java.lang.Object parameter.
   */

  @DISPID(0)
  @DefaultMethod
  com.exceljava.com4j.excel.Series _Default(
    java.lang.Object index);


  // Properties:
}
