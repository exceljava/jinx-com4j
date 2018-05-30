package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020888-0001-0000-C000-000000000046}")
public interface IListBoxes extends Com4jObject,Iterable<Com4jObject> {
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
   */

  @VTID(10)
  void _Dummy3();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(11)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object bringToFront();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(12)
  @ReturnValue(type=NativeType.VARIANT)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {com.exceljava.com4j.excel.XlPictureAppearance.class, com.exceljava.com4j.excel.XlCopyPictureFormat.class}, nativeType = {NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"2", "-4147"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {com.exceljava.com4j.excel.XlCopyPictureFormat.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"-4147"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
  java.lang.Object copyPicture(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlPictureAppearance appearance);

  /**
   * @param appearance Optional parameter. Default value is 2
   * @param format Optional parameter. Default value is -4147
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(13)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object copyPicture(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlPictureAppearance appearance,
    @Optional @DefaultValue("-4147") com.exceljava.com4j.excel.XlCopyPictureFormat format);


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(14)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object cut();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(15)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object delete();


  /**
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(16)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject duplicate();


  /**
   * <p>
   * Getter method for the COM property "Enabled"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(17)
  boolean getEnabled();


  /**
   * <p>
   * Setter method for the COM property "Enabled"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(18)
  void setEnabled(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Height"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(19)
  double getHeight();


  /**
   * <p>
   * Setter method for the COM property "Height"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(20)
  void setHeight(
    double rhs);


  /**
   */

  @VTID(21)
  void _Dummy12();


  /**
   * <p>
   * Getter method for the COM property "Left"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(22)
  double getLeft();


  /**
   * <p>
   * Setter method for the COM property "Left"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(23)
  void setLeft(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Locked"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(24)
  boolean getLocked();


  /**
   * <p>
   * Setter method for the COM property "Locked"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(25)
  void setLocked(
    boolean rhs);


  /**
   */

  @VTID(26)
  void _Dummy15();


  /**
   * <p>
   * Getter method for the COM property "OnAction"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(27)
  java.lang.String getOnAction();


  /**
   * <p>
   * Setter method for the COM property "OnAction"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(28)
  void setOnAction(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Placement"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(29)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getPlacement();


  /**
   * <p>
   * Setter method for the COM property "Placement"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(30)
  void setPlacement(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "PrintObject"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(31)
  boolean getPrintObject();


  /**
   * <p>
   * Setter method for the COM property "PrintObject"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(32)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(33)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object select();

  /**
   * @param replace Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(33)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object select(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object replace);


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(34)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object sendToBack();


  /**
   * <p>
   * Getter method for the COM property "Top"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(35)
  double getTop();


  /**
   * <p>
   * Setter method for the COM property "Top"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(36)
  void setTop(
    double rhs);


  /**
   */

  @VTID(37)
  void _Dummy22();


  /**
   * <p>
   * Getter method for the COM property "Visible"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(38)
  boolean getVisible();


  /**
   * <p>
   * Setter method for the COM property "Visible"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(39)
  void setVisible(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Width"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(40)
  double getWidth();


  /**
   * <p>
   * Setter method for the COM property "Width"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(41)
  void setWidth(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "ZOrder"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(42)
  int getZOrder();


  /**
   * <p>
   * Getter method for the COM property "ShapeRange"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ShapeRange
   */

  @VTID(43)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(44)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
  java.lang.Object addItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object text);

  /**
   * @param text Mandatory java.lang.Object parameter.
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(44)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object addItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object text,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "Display3DShading"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(45)
  boolean getDisplay3DShading();


  /**
   * <p>
   * Setter method for the COM property "Display3DShading"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(46)
  void setDisplay3DShading(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(47)
  @DefaultMethod
  int get_Default();


  /**
   * <p>
   * Setter method for the COM property "_Default"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(48)
  @DefaultMethod
  void set_Default(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "LinkedCell"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(49)
  java.lang.String getLinkedCell();


  /**
   * <p>
   * Setter method for the COM property "LinkedCell"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(50)
  void setLinkedCell(
    java.lang.String rhs);


  /**
   */

  @VTID(51)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(52)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object getList();

  /**
   * <p>
   * Getter method for the COM property "List"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(52)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getList(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);


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

  @VTID(53)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void setList(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "List"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(53)
  void setList(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   */

  @VTID(54)
  void _Dummy33();


  /**
   * <p>
   * Getter method for the COM property "ListFillRange"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(55)
  java.lang.String getListFillRange();


  /**
   * <p>
   * Setter method for the COM property "ListFillRange"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(56)
  void setListFillRange(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "ListIndex"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(57)
  int getListIndex();


  /**
   * <p>
   * Setter method for the COM property "ListIndex"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(58)
  void setListIndex(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "MultiSelect"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(59)
  int getMultiSelect();


  /**
   * <p>
   * Setter method for the COM property "MultiSelect"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(60)
  void setMultiSelect(
    int rhs);


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(61)
  @ReturnValue(type=NativeType.VARIANT)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(62)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
  java.lang.Object removeItem(
    int index);

  /**
   * @param index Mandatory int parameter.
   * @param count Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(62)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object removeItem(
    int index,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object count);


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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(63)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object getSelected();

  /**
   * <p>
   * Getter method for the COM property "Selected"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(63)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getSelected(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);


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

  @VTID(64)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void setSelected(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "Selected"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(64)
  void setSelected(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Value"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(65)
  int getValue();


  /**
   * <p>
   * Setter method for the COM property "Value"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(66)
  void setValue(
    int rhs);


  /**
   * @param left Mandatory double parameter.
   * @param top Mandatory double parameter.
   * @param width Mandatory double parameter.
   * @param height Mandatory double parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.ListBox
   */

  @VTID(67)
  com.exceljava.com4j.excel.ListBox add(
    double left,
    double top,
    double width,
    double height);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(68)
  int getCount();


  /**
   * @return  Returns a value of type com.exceljava.com4j.excel.GroupObject
   */

  @VTID(69)
  com.exceljava.com4j.excel.GroupObject group();


  /**
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(70)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject item(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   */

  @VTID(71)
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
