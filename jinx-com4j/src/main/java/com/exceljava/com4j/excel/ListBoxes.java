package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface ListBoxes extends Com4jObject,Iterable<Com4jObject> {
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
   */

  @DISPID(65539)
  void _Dummy3();


  /**
   */

  @DISPID(602)
  java.lang.Object bringToFront();


  /**
   */

  @DISPID(551)
  java.lang.Object copy();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlPictureAppearance parameter appearance is set to 2</li><li>com.exceljava.com4j.excel.XlCopyPictureFormat parameter format is set to -4147</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * copyPicture(2, -4147);
   * </code>
   * </p>
   */

  @DISPID(213)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {com.exceljava.com4j.excel.XlPictureAppearance.class, com.exceljava.com4j.excel.XlCopyPictureFormat.class}, nativeType = {NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"2", "-4147"})
  @ReturnValue(index=-1)
  java.lang.Object copyPicture();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlCopyPictureFormat parameter format is set to -4147</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * copyPicture(appearance, -4147);
   * </code>
   * </p>
   * @param appearance Optional parameter. Default value is 2
   */

  @DISPID(213)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {com.exceljava.com4j.excel.XlCopyPictureFormat.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"-4147"})
  @ReturnValue(index=-1)
  java.lang.Object copyPicture(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlPictureAppearance appearance);

  /**
   * @param appearance Optional parameter. Default value is 2
   * @param format Optional parameter. Default value is -4147
   */

  @DISPID(213)
  java.lang.Object copyPicture(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlPictureAppearance appearance,
    @Optional @DefaultValue("-4147") com.exceljava.com4j.excel.XlCopyPictureFormat format);


  /**
   */

  @DISPID(565)
  java.lang.Object cut();


  /**
   */

  @DISPID(117)
  java.lang.Object delete();


  /**
   */

  @DISPID(1039)
  com4j.Com4jObject duplicate();


  /**
   * <p>
   * Getter method for the COM property "Enabled"
   * </p>
   */

  @DISPID(600)
  @PropGet
  boolean getEnabled();


  /**
   * <p>
   * Setter method for the COM property "Enabled"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(600)
  @PropPut
  void setEnabled(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Height"
   * </p>
   */

  @DISPID(123)
  @PropGet
  double getHeight();


  /**
   * <p>
   * Setter method for the COM property "Height"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(123)
  @PropPut
  void setHeight(
    double rhs);


  /**
   */

  @DISPID(65548)
  void _Dummy12();


  /**
   * <p>
   * Getter method for the COM property "Left"
   * </p>
   */

  @DISPID(127)
  @PropGet
  double getLeft();


  /**
   * <p>
   * Setter method for the COM property "Left"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(127)
  @PropPut
  void setLeft(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Locked"
   * </p>
   */

  @DISPID(269)
  @PropGet
  boolean getLocked();


  /**
   * <p>
   * Setter method for the COM property "Locked"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(269)
  @PropPut
  void setLocked(
    boolean rhs);


  /**
   */

  @DISPID(65551)
  void _Dummy15();


  /**
   * <p>
   * Getter method for the COM property "OnAction"
   * </p>
   */

  @DISPID(596)
  @PropGet
  java.lang.String getOnAction();


  /**
   * <p>
   * Setter method for the COM property "OnAction"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(596)
  @PropPut
  void setOnAction(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Placement"
   * </p>
   */

  @DISPID(617)
  @PropGet
  java.lang.Object getPlacement();


  /**
   * <p>
   * Setter method for the COM property "Placement"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(617)
  @PropPut
  void setPlacement(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "PrintObject"
   * </p>
   */

  @DISPID(618)
  @PropGet
  boolean getPrintObject();


  /**
   * <p>
   * Setter method for the COM property "PrintObject"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(618)
  @PropPut
  void setPrintObject(
    boolean rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter replace is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * select(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(235)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object select();

  /**
   * @param replace Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(235)
  java.lang.Object select(
    @Optional java.lang.Object replace);


  /**
   */

  @DISPID(605)
  java.lang.Object sendToBack();


  /**
   * <p>
   * Getter method for the COM property "Top"
   * </p>
   */

  @DISPID(126)
  @PropGet
  double getTop();


  /**
   * <p>
   * Setter method for the COM property "Top"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(126)
  @PropPut
  void setTop(
    double rhs);


  /**
   */

  @DISPID(65558)
  void _Dummy22();


  /**
   * <p>
   * Getter method for the COM property "Visible"
   * </p>
   */

  @DISPID(558)
  @PropGet
  boolean getVisible();


  /**
   * <p>
   * Setter method for the COM property "Visible"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(558)
  @PropPut
  void setVisible(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Width"
   * </p>
   */

  @DISPID(122)
  @PropGet
  double getWidth();


  /**
   * <p>
   * Setter method for the COM property "Width"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(122)
  @PropPut
  void setWidth(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "ZOrder"
   * </p>
   */

  @DISPID(622)
  @PropGet
  int getZOrder();


  /**
   * <p>
   * Getter method for the COM property "ShapeRange"
   * </p>
   */

  @DISPID(1528)
  @PropGet
  com.exceljava.com4j.excel.ShapeRange getShapeRange();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addItem(text, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param text Mandatory java.lang.Object parameter.
   */

  @DISPID(851)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object addItem(
    java.lang.Object text);

  /**
   * @param text Mandatory java.lang.Object parameter.
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(851)
  java.lang.Object addItem(
    java.lang.Object text,
    @Optional java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "Display3DShading"
   * </p>
   */

  @DISPID(1122)
  @PropGet
  boolean getDisplay3DShading();


  /**
   * <p>
   * Setter method for the COM property "Display3DShading"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1122)
  @PropPut
  void setDisplay3DShading(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   */

  @DISPID(0)
  @PropGet
  @DefaultMethod
  int get_Default();


  /**
   * <p>
   * Setter method for the COM property "_Default"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(0)
  @PropPut
  @DefaultMethod
  void set_Default(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "LinkedCell"
   * </p>
   */

  @DISPID(1058)
  @PropGet
  java.lang.String getLinkedCell();


  /**
   * <p>
   * Setter method for the COM property "LinkedCell"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1058)
  @PropPut
  void setLinkedCell(
    java.lang.String rhs);


  /**
   */

  @DISPID(65567)
  void _Dummy31();


  /**
   * <p>
   * Getter method for the COM property "List"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getList(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(861)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object getList();

  /**
   * <p>
   * Getter method for the COM property "List"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(861)
  @PropGet
  java.lang.Object getList(
    @Optional java.lang.Object index);


  /**
   * <p>
   * Setter method for the COM property "List"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setList(com4j.Variant.getMissing(), rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(861)
  @PropPut
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void setList(
    java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "List"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(861)
  @PropPut
  void setList(
    @Optional java.lang.Object index,
    java.lang.Object rhs);


  /**
   */

  @DISPID(65569)
  void _Dummy33();


  /**
   * <p>
   * Getter method for the COM property "ListFillRange"
   * </p>
   */

  @DISPID(847)
  @PropGet
  java.lang.String getListFillRange();


  /**
   * <p>
   * Setter method for the COM property "ListFillRange"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(847)
  @PropPut
  void setListFillRange(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "ListIndex"
   * </p>
   */

  @DISPID(850)
  @PropGet
  int getListIndex();


  /**
   * <p>
   * Setter method for the COM property "ListIndex"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(850)
  @PropPut
  void setListIndex(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "MultiSelect"
   * </p>
   */

  @DISPID(32)
  @PropGet
  int getMultiSelect();


  /**
   * <p>
   * Setter method for the COM property "MultiSelect"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(32)
  @PropPut
  void setMultiSelect(
    int rhs);


  /**
   */

  @DISPID(853)
  java.lang.Object removeAllItems();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter count is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * removeItem(index, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param index Mandatory int parameter.
   */

  @DISPID(852)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object removeItem(
    int index);

  /**
   * @param index Mandatory int parameter.
   * @param count Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(852)
  java.lang.Object removeItem(
    int index,
    @Optional java.lang.Object count);


  /**
   * <p>
   * Getter method for the COM property "Selected"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getSelected(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(1123)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object getSelected();

  /**
   * <p>
   * Getter method for the COM property "Selected"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1123)
  @PropGet
  java.lang.Object getSelected(
    @Optional java.lang.Object index);


  /**
   * <p>
   * Setter method for the COM property "Selected"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setSelected(com4j.Variant.getMissing(), rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(1123)
  @PropPut
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void setSelected(
    java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "Selected"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(1123)
  @PropPut
  void setSelected(
    @Optional java.lang.Object index,
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Value"
   * </p>
   */

  @DISPID(6)
  @PropGet
  int getValue();


  /**
   * <p>
   * Setter method for the COM property "Value"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(6)
  @PropPut
  void setValue(
    int rhs);


  /**
   * @param left Mandatory double parameter.
   * @param top Mandatory double parameter.
   * @param width Mandatory double parameter.
   * @param height Mandatory double parameter.
   */

  @DISPID(181)
  com.exceljava.com4j.excel.ListBox add(
    double left,
    double top,
    double width,
    double height);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   */

  @DISPID(118)
  @PropGet
  int getCount();


  /**
   */

  @DISPID(46)
  com.exceljava.com4j.excel.GroupObject group();


  /**
   * @param index Mandatory java.lang.Object parameter.
   */

  @DISPID(170)
  com4j.Com4jObject item(
    java.lang.Object index);


  /**
   */

  @DISPID(-4)
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
