package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0362-0000-0000-C000-000000000046}")
public interface FileDialog extends com.exceljava.com4j.office._IMsoDispObj {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1610809344) //= 0x60030000. The runtime will prefer the VTID if present
  @VTID(9)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "Filters"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.FileDialogFilters
   */

  @DISPID(1610809345) //= 0x60030001. The runtime will prefer the VTID if present
  @VTID(10)
  com.exceljava.com4j.office.FileDialogFilters getFilters();


  @VTID(10)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.FileDialogFilters.class})
  com.exceljava.com4j.office.FileDialogFilter getFilters(
    int index);

  /**
   * <p>
   * Getter method for the COM property "FilterIndex"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610809346) //= 0x60030002. The runtime will prefer the VTID if present
  @VTID(11)
  int getFilterIndex();


  /**
   * <p>
   * Setter method for the COM property "FilterIndex"
   * </p>
   * @param filterIndex Mandatory int parameter.
   */

  @DISPID(1610809346) //= 0x60030002. The runtime will prefer the VTID if present
  @VTID(12)
  void setFilterIndex(
    int filterIndex);


  /**
   * <p>
   * Getter method for the COM property "Title"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610809348) //= 0x60030004. The runtime will prefer the VTID if present
  @VTID(13)
  java.lang.String getTitle();


  /**
   * <p>
   * Setter method for the COM property "Title"
   * </p>
   * @param title Mandatory java.lang.String parameter.
   */

  @DISPID(1610809348) //= 0x60030004. The runtime will prefer the VTID if present
  @VTID(14)
  void setTitle(
    java.lang.String title);


  /**
   * <p>
   * Getter method for the COM property "ButtonName"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610809350) //= 0x60030006. The runtime will prefer the VTID if present
  @VTID(15)
  java.lang.String getButtonName();


  /**
   * <p>
   * Setter method for the COM property "ButtonName"
   * </p>
   * @param buttonName Mandatory java.lang.String parameter.
   */

  @DISPID(1610809350) //= 0x60030006. The runtime will prefer the VTID if present
  @VTID(16)
  void setButtonName(
    java.lang.String buttonName);


  /**
   * <p>
   * Getter method for the COM property "AllowMultiSelect"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610809352) //= 0x60030008. The runtime will prefer the VTID if present
  @VTID(17)
  boolean getAllowMultiSelect();


  /**
   * <p>
   * Setter method for the COM property "AllowMultiSelect"
   * </p>
   * @param pvarfAllowMultiSelect Mandatory boolean parameter.
   */

  @DISPID(1610809352) //= 0x60030008. The runtime will prefer the VTID if present
  @VTID(18)
  void setAllowMultiSelect(
    boolean pvarfAllowMultiSelect);


  /**
   * <p>
   * Getter method for the COM property "InitialView"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoFileDialogView
   */

  @DISPID(1610809354) //= 0x6003000a. The runtime will prefer the VTID if present
  @VTID(19)
  com.exceljava.com4j.office.MsoFileDialogView getInitialView();


  /**
   * <p>
   * Setter method for the COM property "InitialView"
   * </p>
   * @param pinitialview Mandatory com.exceljava.com4j.office.MsoFileDialogView parameter.
   */

  @DISPID(1610809354) //= 0x6003000a. The runtime will prefer the VTID if present
  @VTID(20)
  void setInitialView(
    com.exceljava.com4j.office.MsoFileDialogView pinitialview);


  /**
   * <p>
   * Getter method for the COM property "InitialFileName"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610809356) //= 0x6003000c. The runtime will prefer the VTID if present
  @VTID(21)
  java.lang.String getInitialFileName();


  /**
   * <p>
   * Setter method for the COM property "InitialFileName"
   * </p>
   * @param initialFileName Mandatory java.lang.String parameter.
   */

  @DISPID(1610809356) //= 0x6003000c. The runtime will prefer the VTID if present
  @VTID(22)
  void setInitialFileName(
    java.lang.String initialFileName);


  /**
   * <p>
   * Getter method for the COM property "SelectedItems"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.FileDialogSelectedItems
   */

  @DISPID(1610809358) //= 0x6003000e. The runtime will prefer the VTID if present
  @VTID(23)
  com.exceljava.com4j.office.FileDialogSelectedItems getSelectedItems();


  @VTID(23)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.FileDialogSelectedItems.class})
  java.lang.String getSelectedItems(
    int index);

  /**
   * <p>
   * Getter method for the COM property "DialogType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoFileDialogType
   */

  @DISPID(1610809359) //= 0x6003000f. The runtime will prefer the VTID if present
  @VTID(24)
  com.exceljava.com4j.office.MsoFileDialogType getDialogType();


  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(25)
  @DefaultMethod
  java.lang.String getItem();


  /**
   * @return  Returns a value of type int
   */

  @DISPID(1610809361) //= 0x60030011. The runtime will prefer the VTID if present
  @VTID(26)
  int show();


  /**
   */

  @DISPID(1610809362) //= 0x60030012. The runtime will prefer the VTID if present
  @VTID(27)
  void execute();


  // Properties:
}
