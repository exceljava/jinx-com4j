package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000CDB06-0000-0000-C000-000000000046}")
public interface ICustomXMLPartEvents extends Com4jObject {
  // Methods:
  /**
   * @param newNode Mandatory com.exceljava.com4j.office.CustomXMLNode parameter.
   * @param inUndoRedo Mandatory boolean parameter.
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(7)
  void nodeAfterInsert(
    com.exceljava.com4j.office.CustomXMLNode newNode,
    boolean inUndoRedo);


  /**
   * @param oldNode Mandatory com.exceljava.com4j.office.CustomXMLNode parameter.
   * @param oldParentNode Mandatory com.exceljava.com4j.office.CustomXMLNode parameter.
   * @param oldNextSibling Mandatory com.exceljava.com4j.office.CustomXMLNode parameter.
   * @param inUndoRedo Mandatory boolean parameter.
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(8)
  void nodeAfterDelete(
    com.exceljava.com4j.office.CustomXMLNode oldNode,
    com.exceljava.com4j.office.CustomXMLNode oldParentNode,
    com.exceljava.com4j.office.CustomXMLNode oldNextSibling,
    boolean inUndoRedo);


  /**
   * @param oldNode Mandatory com.exceljava.com4j.office.CustomXMLNode parameter.
   * @param newNode Mandatory com.exceljava.com4j.office.CustomXMLNode parameter.
   * @param inUndoRedo Mandatory boolean parameter.
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(9)
  void nodeAfterReplace(
    com.exceljava.com4j.office.CustomXMLNode oldNode,
    com.exceljava.com4j.office.CustomXMLNode newNode,
    boolean inUndoRedo);


  // Properties:
}
