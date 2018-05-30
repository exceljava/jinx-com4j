package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024447-0001-0000-C000-000000000046}")
public interface IPhonetics extends Com4jObject,Iterable<Com4jObject> {
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
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(10)
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "Start"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(11)
  int getStart();


  /**
   * <p>
   * Getter method for the COM property "Length"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(12)
  int getLength();


  /**
   * <p>
   * Getter method for the COM property "Visible"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(13)
  boolean getVisible();


  /**
   * <p>
   * Setter method for the COM property "Visible"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(14)
  void setVisible(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "CharacterType"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(15)
  int getCharacterType();


  /**
   * <p>
   * Setter method for the COM property "CharacterType"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(16)
  void setCharacterType(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Alignment"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(17)
  int getAlignment();


  /**
   * <p>
   * Setter method for the COM property "Alignment"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(18)
  void setAlignment(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Font"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Font
   */

  @VTID(19)
  com.exceljava.com4j.excel.Font getFont();


  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(20)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getItem(
    int index);


  /**
   */

  @VTID(21)
  void delete();


  /**
   * @param start Mandatory int parameter.
   * @param length Mandatory int parameter.
   * @param text Mandatory java.lang.String parameter.
   */

  @VTID(22)
  void add(
    int start,
    int length,
    java.lang.String text);


  /**
   * <p>
   * Getter method for the COM property "Text"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(23)
  java.lang.String getText();


  /**
   * <p>
   * Setter method for the COM property "Text"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(24)
  void setText(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(25)
  @DefaultMethod
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject get_Default(
    int index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(26)
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
