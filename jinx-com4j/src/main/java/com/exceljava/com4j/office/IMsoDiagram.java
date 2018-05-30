package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C036D-0000-0000-C000-000000000046}")
public interface IMsoDiagram extends com.exceljava.com4j.office._IMsoDispObj {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(100) //= 0x64. The runtime will prefer the VTID if present
  @VTID(9)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "Nodes"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.DiagramNodes
   */

  @DISPID(101) //= 0x65. The runtime will prefer the VTID if present
  @VTID(10)
  com.exceljava.com4j.office.DiagramNodes getNodes();


  @VTID(10)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.DiagramNodes.class})
  com.exceljava.com4j.excel.DiagramNode getNodes(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoDiagramType
   */

  @DISPID(102) //= 0x66. The runtime will prefer the VTID if present
  @VTID(11)
  com.exceljava.com4j.office.MsoDiagramType getType();


  /**
   * <p>
   * Getter method for the COM property "AutoLayout"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(103) //= 0x67. The runtime will prefer the VTID if present
  @VTID(12)
  com.exceljava.com4j.office.MsoTriState getAutoLayout();


  /**
   * <p>
   * Setter method for the COM property "AutoLayout"
   * </p>
   * @param autoLayout Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(103) //= 0x67. The runtime will prefer the VTID if present
  @VTID(13)
  void setAutoLayout(
    com.exceljava.com4j.office.MsoTriState autoLayout);


  /**
   * <p>
   * Getter method for the COM property "Reverse"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(104) //= 0x68. The runtime will prefer the VTID if present
  @VTID(14)
  com.exceljava.com4j.office.MsoTriState getReverse();


  /**
   * <p>
   * Setter method for the COM property "Reverse"
   * </p>
   * @param reverse Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(104) //= 0x68. The runtime will prefer the VTID if present
  @VTID(15)
  void setReverse(
    com.exceljava.com4j.office.MsoTriState reverse);


  /**
   * <p>
   * Getter method for the COM property "AutoFormat"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(105) //= 0x69. The runtime will prefer the VTID if present
  @VTID(16)
  com.exceljava.com4j.office.MsoTriState getAutoFormat();


  /**
   * <p>
   * Setter method for the COM property "AutoFormat"
   * </p>
   * @param autoFormat Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(105) //= 0x69. The runtime will prefer the VTID if present
  @VTID(17)
  void setAutoFormat(
    com.exceljava.com4j.office.MsoTriState autoFormat);


  /**
   * @param type Mandatory com.exceljava.com4j.office.MsoDiagramType parameter.
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(18)
  void convert(
    com.exceljava.com4j.office.MsoDiagramType type);


  /**
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(19)
  void fitText();


  // Properties:
}
