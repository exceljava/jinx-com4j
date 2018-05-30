package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020893-0001-0000-C000-000000000046}")
public interface IWindow extends Com4jObject {
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(10)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object activate();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(11)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object activateNext();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(12)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object activatePrevious();


  /**
   * <p>
   * Getter method for the COM property "ActiveCell"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(13)
  com.exceljava.com4j.excel.Range getActiveCell();


  /**
   * <p>
   * Getter method for the COM property "ActiveChart"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel._Chart
   */

  @VTID(14)
  com.exceljava.com4j.excel._Chart getActiveChart();


  /**
   * <p>
   * Getter method for the COM property "ActivePane"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Pane
   */

  @VTID(15)
  com.exceljava.com4j.excel.Pane getActivePane();


  /**
   * <p>
   * Getter method for the COM property "ActiveSheet"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(16)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getActiveSheet();


  /**
   * <p>
   * Getter method for the COM property "Caption"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(17)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getCaption();


  /**
   * <p>
   * Setter method for the COM property "Caption"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(18)
  void setCaption(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter saveChanges is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter filename is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter routeWorkbook is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * close(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(19)
  @UseDefaultValues(paramIndexMapping = {3}, optParamIndex = {0, 1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=3)
  boolean close();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter filename is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter routeWorkbook is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * close(saveChanges, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param saveChanges Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(19)
  @UseDefaultValues(paramIndexMapping = {0, 3}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=3)
  boolean close(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveChanges);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter routeWorkbook is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * close(saveChanges, filename, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param saveChanges Optional parameter. Default value is com4j.Variant.getMissing()
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(19)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=3)
  boolean close(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveChanges,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename);

  /**
   * @param saveChanges Optional parameter. Default value is com4j.Variant.getMissing()
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param routeWorkbook Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(19)
  boolean close(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveChanges,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object routeWorkbook);


  /**
   * <p>
   * Getter method for the COM property "DisplayFormulas"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(20)
  boolean getDisplayFormulas();


  /**
   * <p>
   * Setter method for the COM property "DisplayFormulas"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(21)
  void setDisplayFormulas(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DisplayGridlines"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(22)
  boolean getDisplayGridlines();


  /**
   * <p>
   * Setter method for the COM property "DisplayGridlines"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(23)
  void setDisplayGridlines(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DisplayHeadings"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(24)
  boolean getDisplayHeadings();


  /**
   * <p>
   * Setter method for the COM property "DisplayHeadings"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(25)
  void setDisplayHeadings(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DisplayHorizontalScrollBar"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(26)
  boolean getDisplayHorizontalScrollBar();


  /**
   * <p>
   * Setter method for the COM property "DisplayHorizontalScrollBar"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(27)
  void setDisplayHorizontalScrollBar(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DisplayOutline"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(28)
  boolean getDisplayOutline();


  /**
   * <p>
   * Setter method for the COM property "DisplayOutline"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(29)
  void setDisplayOutline(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "_DisplayRightToLeft"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(30)
  boolean get_DisplayRightToLeft();


  /**
   * <p>
   * Setter method for the COM property "_DisplayRightToLeft"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(31)
  void set_DisplayRightToLeft(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DisplayVerticalScrollBar"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(32)
  boolean getDisplayVerticalScrollBar();


  /**
   * <p>
   * Setter method for the COM property "DisplayVerticalScrollBar"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(33)
  void setDisplayVerticalScrollBar(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DisplayWorkbookTabs"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(34)
  boolean getDisplayWorkbookTabs();


  /**
   * <p>
   * Setter method for the COM property "DisplayWorkbookTabs"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(35)
  void setDisplayWorkbookTabs(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DisplayZeros"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(36)
  boolean getDisplayZeros();


  /**
   * <p>
   * Setter method for the COM property "DisplayZeros"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(37)
  void setDisplayZeros(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "EnableResize"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(38)
  boolean getEnableResize();


  /**
   * <p>
   * Setter method for the COM property "EnableResize"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(39)
  void setEnableResize(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "FreezePanes"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(40)
  boolean getFreezePanes();


  /**
   * <p>
   * Setter method for the COM property "FreezePanes"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(41)
  void setFreezePanes(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "GridlineColor"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(42)
  int getGridlineColor();


  /**
   * <p>
   * Setter method for the COM property "GridlineColor"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(43)
  void setGridlineColor(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "GridlineColorIndex"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlColorIndex
   */

  @VTID(44)
  com.exceljava.com4j.excel.XlColorIndex getGridlineColorIndex();


  /**
   * <p>
   * Setter method for the COM property "GridlineColorIndex"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlColorIndex parameter.
   */

  @VTID(45)
  void setGridlineColorIndex(
    com.exceljava.com4j.excel.XlColorIndex rhs);


  /**
   * <p>
   * Getter method for the COM property "Height"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(46)
  double getHeight();


  /**
   * <p>
   * Setter method for the COM property "Height"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(47)
  void setHeight(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Index"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(48)
  int getIndex();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter down is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter up is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter toRight is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter toLeft is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * largeScroll(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(49)
  @UseDefaultValues(paramIndexMapping = {4}, optParamIndex = {0, 1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object largeScroll();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter up is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter toRight is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter toLeft is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * largeScroll(down, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param down Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(49)
  @UseDefaultValues(paramIndexMapping = {0, 4}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object largeScroll(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object down);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter toRight is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter toLeft is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * largeScroll(down, up, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param down Optional parameter. Default value is com4j.Variant.getMissing()
   * @param up Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(49)
  @UseDefaultValues(paramIndexMapping = {0, 1, 4}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object largeScroll(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object down,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object up);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter toLeft is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * largeScroll(down, up, toRight, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param down Optional parameter. Default value is com4j.Variant.getMissing()
   * @param up Optional parameter. Default value is com4j.Variant.getMissing()
   * @param toRight Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(49)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 4}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object largeScroll(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object down,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object up,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object toRight);

  /**
   * @param down Optional parameter. Default value is com4j.Variant.getMissing()
   * @param up Optional parameter. Default value is com4j.Variant.getMissing()
   * @param toRight Optional parameter. Default value is com4j.Variant.getMissing()
   * @param toLeft Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(49)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object largeScroll(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object down,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object up,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object toRight,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object toLeft);


  /**
   * <p>
   * Getter method for the COM property "Left"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(50)
  double getLeft();


  /**
   * <p>
   * Setter method for the COM property "Left"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(51)
  void setLeft(
    double rhs);


  /**
   * @return  Returns a value of type com.exceljava.com4j.excel.Window
   */

  @VTID(52)
  com.exceljava.com4j.excel.Window newWindow();


  /**
   * <p>
   * Getter method for the COM property "OnWindow"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(53)
  java.lang.String getOnWindow();


  /**
   * <p>
   * Setter method for the COM property "OnWindow"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(54)
  void setOnWindow(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Panes"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Panes
   */

  @VTID(55)
  com.exceljava.com4j.excel.Panes getPanes();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(56)
  @UseDefaultValues(paramIndexMapping = {8}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object _PrintOut();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(56)
  @UseDefaultValues(paramIndexMapping = {0, 8}, optParamIndex = {1, 2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, to, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(56)
  @UseDefaultValues(paramIndexMapping = {0, 1, 8}, optParamIndex = {2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, to, copies, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(56)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 8}, optParamIndex = {3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, to, copies, preview, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(56)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 8}, optParamIndex = {4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, to, copies, preview, activePrinter, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(56)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 8}, optParamIndex = {5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, to, copies, preview, activePrinter, printToFile, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(56)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 8}, optParamIndex = {6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, to, copies, preview, activePrinter, printToFile, collate, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param collate Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(56)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 8}, optParamIndex = {7}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object collate);

  /**
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param collate Optional parameter. Default value is com4j.Variant.getMissing()
   * @param prToFileName Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(56)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object collate,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object prToFileName);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter enableChanges is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printPreview(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(57)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object printPreview();

  /**
   * @param enableChanges Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(57)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object printPreview(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object enableChanges);


  /**
   * <p>
   * Getter method for the COM property "RangeSelection"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(58)
  com.exceljava.com4j.excel.Range getRangeSelection();


  /**
   * <p>
   * Getter method for the COM property "ScrollColumn"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(59)
  int getScrollColumn();


  /**
   * <p>
   * Setter method for the COM property "ScrollColumn"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(60)
  void setScrollColumn(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "ScrollRow"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(61)
  int getScrollRow();


  /**
   * <p>
   * Setter method for the COM property "ScrollRow"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(62)
  void setScrollRow(
    int rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter sheets is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter position is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * scrollWorkbookTabs(com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(63)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
  java.lang.Object scrollWorkbookTabs();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter position is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * scrollWorkbookTabs(sheets, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param sheets Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(63)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
  java.lang.Object scrollWorkbookTabs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sheets);

  /**
   * @param sheets Optional parameter. Default value is com4j.Variant.getMissing()
   * @param position Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(63)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object scrollWorkbookTabs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sheets,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object position);


  /**
   * <p>
   * Getter method for the COM property "SelectedSheets"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Sheets
   */

  @VTID(64)
  com.exceljava.com4j.excel.Sheets getSelectedSheets();


  @VTID(64)
  @ReturnValue(type=NativeType.Dispatch,defaultPropertyThrough={com.exceljava.com4j.excel.Sheets.class})
  com4j.Com4jObject getSelectedSheets(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "Selection"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(65)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getSelection();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter down is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter up is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter toRight is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter toLeft is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * smallScroll(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(66)
  @UseDefaultValues(paramIndexMapping = {4}, optParamIndex = {0, 1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object smallScroll();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter up is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter toRight is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter toLeft is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * smallScroll(down, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param down Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(66)
  @UseDefaultValues(paramIndexMapping = {0, 4}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object smallScroll(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object down);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter toRight is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter toLeft is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * smallScroll(down, up, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param down Optional parameter. Default value is com4j.Variant.getMissing()
   * @param up Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(66)
  @UseDefaultValues(paramIndexMapping = {0, 1, 4}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object smallScroll(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object down,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object up);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter toLeft is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * smallScroll(down, up, toRight, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param down Optional parameter. Default value is com4j.Variant.getMissing()
   * @param up Optional parameter. Default value is com4j.Variant.getMissing()
   * @param toRight Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(66)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 4}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object smallScroll(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object down,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object up,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object toRight);

  /**
   * @param down Optional parameter. Default value is com4j.Variant.getMissing()
   * @param up Optional parameter. Default value is com4j.Variant.getMissing()
   * @param toRight Optional parameter. Default value is com4j.Variant.getMissing()
   * @param toLeft Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(66)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object smallScroll(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object down,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object up,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object toRight,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object toLeft);


  /**
   * <p>
   * Getter method for the COM property "Split"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(67)
  boolean getSplit();


  /**
   * <p>
   * Setter method for the COM property "Split"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(68)
  void setSplit(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "SplitColumn"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(69)
  int getSplitColumn();


  /**
   * <p>
   * Setter method for the COM property "SplitColumn"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(70)
  void setSplitColumn(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "SplitHorizontal"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(71)
  double getSplitHorizontal();


  /**
   * <p>
   * Setter method for the COM property "SplitHorizontal"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(72)
  void setSplitHorizontal(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "SplitRow"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(73)
  int getSplitRow();


  /**
   * <p>
   * Setter method for the COM property "SplitRow"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(74)
  void setSplitRow(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "SplitVertical"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(75)
  double getSplitVertical();


  /**
   * <p>
   * Setter method for the COM property "SplitVertical"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(76)
  void setSplitVertical(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "TabRatio"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(77)
  double getTabRatio();


  /**
   * <p>
   * Setter method for the COM property "TabRatio"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(78)
  void setTabRatio(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Top"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(79)
  double getTop();


  /**
   * <p>
   * Setter method for the COM property "Top"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(80)
  void setTop(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlWindowType
   */

  @VTID(81)
  com.exceljava.com4j.excel.XlWindowType getType();


  /**
   * <p>
   * Getter method for the COM property "UsableHeight"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(82)
  double getUsableHeight();


  /**
   * <p>
   * Getter method for the COM property "UsableWidth"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(83)
  double getUsableWidth();


  /**
   * <p>
   * Getter method for the COM property "Visible"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(84)
  boolean getVisible();


  /**
   * <p>
   * Setter method for the COM property "Visible"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(85)
  void setVisible(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "VisibleRange"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(86)
  com.exceljava.com4j.excel.Range getVisibleRange();


  /**
   * <p>
   * Getter method for the COM property "Width"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(87)
  double getWidth();


  /**
   * <p>
   * Setter method for the COM property "Width"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(88)
  void setWidth(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "WindowNumber"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(89)
  int getWindowNumber();


  /**
   * <p>
   * Getter method for the COM property "WindowState"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlWindowState
   */

  @VTID(90)
  com.exceljava.com4j.excel.XlWindowState getWindowState();


  /**
   * <p>
   * Setter method for the COM property "WindowState"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlWindowState parameter.
   */

  @VTID(91)
  void setWindowState(
    com.exceljava.com4j.excel.XlWindowState rhs);


  /**
   * <p>
   * Getter method for the COM property "Zoom"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(92)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getZoom();


  /**
   * <p>
   * Setter method for the COM property "Zoom"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(93)
  void setZoom(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "View"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlWindowView
   */

  @VTID(94)
  com.exceljava.com4j.excel.XlWindowView getView();


  /**
   * <p>
   * Setter method for the COM property "View"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlWindowView parameter.
   */

  @VTID(95)
  void setView(
    com.exceljava.com4j.excel.XlWindowView rhs);


  /**
   * <p>
   * Getter method for the COM property "DisplayRightToLeft"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(96)
  boolean getDisplayRightToLeft();


  /**
   * <p>
   * Setter method for the COM property "DisplayRightToLeft"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(97)
  void setDisplayRightToLeft(
    boolean rhs);


  /**
   * @param points Mandatory int parameter.
   * @return  Returns a value of type int
   */

  @VTID(98)
  int pointsToScreenPixelsX(
    int points);


  /**
   * @param points Mandatory int parameter.
   * @return  Returns a value of type int
   */

  @VTID(99)
  int pointsToScreenPixelsY(
    int points);


  /**
   * @param x Mandatory int parameter.
   * @param y Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(100)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject rangeFromPoint(
    int x,
    int y);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter start is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * scrollIntoView(left, top, width, height, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param left Mandatory int parameter.
   * @param top Mandatory int parameter.
   * @param width Mandatory int parameter.
   * @param height Mandatory int parameter.
   */

  @VTID(101)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void scrollIntoView(
    int left,
    int top,
    int width,
    int height);

  /**
   * @param left Mandatory int parameter.
   * @param top Mandatory int parameter.
   * @param width Mandatory int parameter.
   * @param height Mandatory int parameter.
   * @param start Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(101)
  void scrollIntoView(
    int left,
    int top,
    int width,
    int height,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object start);


  /**
   * <p>
   * Getter method for the COM property "SheetViews"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.SheetViews
   */

  @VTID(102)
  com.exceljava.com4j.excel.SheetViews getSheetViews();


  /**
   * <p>
   * Getter method for the COM property "ActiveSheetView"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(103)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getActiveSheetView();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(104)
  @UseDefaultValues(paramIndexMapping = {8}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object printOut();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(104)
  @UseDefaultValues(paramIndexMapping = {0, 8}, optParamIndex = {1, 2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(104)
  @UseDefaultValues(paramIndexMapping = {0, 1, 8}, optParamIndex = {2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(104)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 8}, optParamIndex = {3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, preview, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(104)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 8}, optParamIndex = {4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, preview, activePrinter, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(104)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 8}, optParamIndex = {5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, preview, activePrinter, printToFile, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(104)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 8}, optParamIndex = {6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, preview, activePrinter, printToFile, collate, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param collate Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(104)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 8}, optParamIndex = {7}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=8)
  java.lang.Object printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object collate);

  /**
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param collate Optional parameter. Default value is com4j.Variant.getMissing()
   * @param prToFileName Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(104)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object collate,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object prToFileName);


  /**
   * <p>
   * Getter method for the COM property "DisplayRuler"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(105)
  boolean getDisplayRuler();


  /**
   * <p>
   * Setter method for the COM property "DisplayRuler"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(106)
  void setDisplayRuler(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "AutoFilterDateGrouping"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(107)
  boolean getAutoFilterDateGrouping();


  /**
   * <p>
   * Setter method for the COM property "AutoFilterDateGrouping"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(108)
  void setAutoFilterDateGrouping(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DisplayWhitespace"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(109)
  boolean getDisplayWhitespace();


  /**
   * <p>
   * Setter method for the COM property "DisplayWhitespace"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(110)
  void setDisplayWhitespace(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Hwnd"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(111)
  int getHwnd();


  // Properties:
}
