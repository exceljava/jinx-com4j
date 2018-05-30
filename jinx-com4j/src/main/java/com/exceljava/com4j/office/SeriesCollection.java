package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C170A-0000-0000-C000-000000000046}")
public interface SeriesCollection extends Com4jObject,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(7)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.XlRowCol parameter rowcol is set to 2</li><li>java.lang.Object parameter seriesLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter categoryLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter replace is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(source, 2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param source Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoSeries
   */

  @VTID(8)
  @UseDefaultValues(paramIndexMapping = {0, 5}, optParamIndex = {1, 2, 3, 4}, javaType = {com.exceljava.com4j.office.XlRowCol.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"2", "80020004", "80020004", "80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.office.IMsoSeries add(
    @MarshalAs(NativeType.VARIANT) java.lang.Object source);

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
   * @param rowcol Optional parameter. Default value is 2
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoSeries
   */

  @VTID(8)
  @UseDefaultValues(paramIndexMapping = {0, 1, 5}, optParamIndex = {2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.office.IMsoSeries add(
    @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @Optional @DefaultValue("2") com.exceljava.com4j.office.XlRowCol rowcol);

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
   * @param rowcol Optional parameter. Default value is 2
   * @param seriesLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoSeries
   */

  @VTID(8)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 5}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.office.IMsoSeries add(
    @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @Optional @DefaultValue("2") com.exceljava.com4j.office.XlRowCol rowcol,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object seriesLabels);

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
   * @param rowcol Optional parameter. Default value is 2
   * @param seriesLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param categoryLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoSeries
   */

  @VTID(8)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 5}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.office.IMsoSeries add(
    @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @Optional @DefaultValue("2") com.exceljava.com4j.office.XlRowCol rowcol,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object seriesLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object categoryLabels);

  /**
   * @param source Mandatory java.lang.Object parameter.
   * @param rowcol Optional parameter. Default value is 2
   * @param seriesLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param categoryLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param replace Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoSeries
   */

  @VTID(8)
  com.exceljava.com4j.office.IMsoSeries add(
    @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @Optional @DefaultValue("2") com.exceljava.com4j.office.XlRowCol rowcol,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object seriesLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object categoryLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object replace);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(9)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 3}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=3)
  java.lang.Object extend(
    @MarshalAs(NativeType.VARIANT) java.lang.Object source);

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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=3)
  java.lang.Object extend(
    @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowcol);

  /**
   * @param source Mandatory java.lang.Object parameter.
   * @param rowcol Optional parameter. Default value is com4j.Variant.getMissing()
   * @param categoryLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(10)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object extend(
    @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowcol,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object categoryLabels);


  /**
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoSeries
   */

  @VTID(11)
  com.exceljava.com4j.office.IMsoSeries item(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   */

  @VTID(12)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.XlRowCol parameter rowcol is set to 2</li><li>java.lang.Object parameter seriesLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter categoryLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter replace is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter newSeries is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * paste(2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {5}, optParamIndex = {0, 1, 2, 3, 4}, javaType = {com.exceljava.com4j.office.XlRowCol.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"2", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=5)
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
   * @param rowcol Optional parameter. Default value is 2
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {0, 5}, optParamIndex = {1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=5)
  java.lang.Object paste(
    @Optional @DefaultValue("2") com.exceljava.com4j.office.XlRowCol rowcol);

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
   * @param rowcol Optional parameter. Default value is 2
   * @param seriesLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {0, 1, 5}, optParamIndex = {2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=5)
  java.lang.Object paste(
    @Optional @DefaultValue("2") com.exceljava.com4j.office.XlRowCol rowcol,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object seriesLabels);

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
   * @param rowcol Optional parameter. Default value is 2
   * @param seriesLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param categoryLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 5}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=5)
  java.lang.Object paste(
    @Optional @DefaultValue("2") com.exceljava.com4j.office.XlRowCol rowcol,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object seriesLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object categoryLabels);

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
   * @param rowcol Optional parameter. Default value is 2
   * @param seriesLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param categoryLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param replace Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 5}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=5)
  java.lang.Object paste(
    @Optional @DefaultValue("2") com.exceljava.com4j.office.XlRowCol rowcol,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object seriesLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object categoryLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object replace);

  /**
   * @param rowcol Optional parameter. Default value is 2
   * @param seriesLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param categoryLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param replace Optional parameter. Default value is com4j.Variant.getMissing()
   * @param newSeries Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(13)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object paste(
    @Optional @DefaultValue("2") com.exceljava.com4j.office.XlRowCol rowcol,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object seriesLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object categoryLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object replace,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object newSeries);


  /**
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoSeries
   */

  @VTID(14)
  com.exceljava.com4j.office.IMsoSeries newSeries();


  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(15)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(16)
  int getCreator();


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoSeries
   */

  @VTID(17)
  @DefaultMethod
  com.exceljava.com4j.office.IMsoSeries get_Default(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  // Properties:
}
