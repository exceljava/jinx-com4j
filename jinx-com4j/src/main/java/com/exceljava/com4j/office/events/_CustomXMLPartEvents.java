package com.exceljava.com4j.office.events;

import com4j.*;

@IID("{000CDB07-0000-0000-C000-000000000046}")
public abstract class _CustomXMLPartEvents {
  // Methods:
  /**
   * @param newNode Mandatory com.exceljava.com4j.office.CustomXMLNode parameter.
   * @param inUndoRedo Mandatory boolean parameter.
   */

  @DISPID(1)
  public void nodeAfterInsert(
    com.exceljava.com4j.office.CustomXMLNode newNode,
    boolean inUndoRedo) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param oldNode Mandatory com.exceljava.com4j.office.CustomXMLNode parameter.
   * @param oldParentNode Mandatory com.exceljava.com4j.office.CustomXMLNode parameter.
   * @param oldNextSibling Mandatory com.exceljava.com4j.office.CustomXMLNode parameter.
   * @param inUndoRedo Mandatory boolean parameter.
   */

  @DISPID(2)
  public void nodeAfterDelete(
    com.exceljava.com4j.office.CustomXMLNode oldNode,
    com.exceljava.com4j.office.CustomXMLNode oldParentNode,
    com.exceljava.com4j.office.CustomXMLNode oldNextSibling,
    boolean inUndoRedo) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param oldNode Mandatory com.exceljava.com4j.office.CustomXMLNode parameter.
   * @param newNode Mandatory com.exceljava.com4j.office.CustomXMLNode parameter.
   * @param inUndoRedo Mandatory boolean parameter.
   */

  @DISPID(3)
  public void nodeAfterReplace(
    com.exceljava.com4j.office.CustomXMLNode oldNode,
    com.exceljava.com4j.office.CustomXMLNode newNode,
    boolean inUndoRedo) {
        throw new UnsupportedOperationException();
  }


  // Properties:
}
