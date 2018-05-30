package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002446F-0001-0000-C000-000000000046}")
public interface IDiagram extends Com4jObject {
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
   * Getter method for the COM property "Nodes"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.DiagramNodes
   */

  @VTID(10)
  com.exceljava.com4j.excel.DiagramNodes getNodes();


  @VTID(10)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.excel.DiagramNodes.class})
  com.exceljava.com4j.excel.DiagramNode getNodes(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoDiagramType
   */

  @VTID(11)
  com.exceljava.com4j.office.MsoDiagramType getType();


  /**
   * <p>
   * Getter method for the COM property "AutoLayout"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @VTID(12)
  com.exceljava.com4j.office.MsoTriState getAutoLayout();


  /**
   * <p>
   * Setter method for the COM property "AutoLayout"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @VTID(13)
  void setAutoLayout(
    com.exceljava.com4j.office.MsoTriState rhs);


  /**
   * <p>
   * Getter method for the COM property "Reverse"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @VTID(14)
  com.exceljava.com4j.office.MsoTriState getReverse();


  /**
   * <p>
   * Setter method for the COM property "Reverse"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @VTID(15)
  void setReverse(
    com.exceljava.com4j.office.MsoTriState rhs);


  /**
   * <p>
   * Getter method for the COM property "AutoFormat"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @VTID(16)
  com.exceljava.com4j.office.MsoTriState getAutoFormat();


  /**
   * <p>
   * Setter method for the COM property "AutoFormat"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @VTID(17)
  void setAutoFormat(
    com.exceljava.com4j.office.MsoTriState rhs);


  /**
   * @param type Mandatory com.exceljava.com4j.office.MsoDiagramType parameter.
   */

  @VTID(18)
  void convert(
    com.exceljava.com4j.office.MsoDiagramType type);


  /**
   */

  @VTID(19)
  void fitText();


  // Properties:
}
