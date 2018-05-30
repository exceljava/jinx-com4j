package com.exceljava.com4j.vbide  ;

import com4j.*;

@IID("{0002E166-0000-0000-C000-000000000046}")
public interface VBE extends com.exceljava.com4j.vbide.Application {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "VBProjects"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide._VBProjects
   */

  @DISPID(107) //= 0x6b. The runtime will prefer the VTID if present
  @VTID(8)
  com.exceljava.com4j.vbide._VBProjects getVBProjects();


  /**
   * <p>
   * Getter method for the COM property "CommandBars"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office._CommandBars
   */

  @DISPID(108) //= 0x6c. The runtime will prefer the VTID if present
  @VTID(9)
  com.exceljava.com4j.office._CommandBars getCommandBars();


  /**
   * <p>
   * Getter method for the COM property "CodePanes"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide._CodePanes
   */

  @DISPID(109) //= 0x6d. The runtime will prefer the VTID if present
  @VTID(10)
  com.exceljava.com4j.vbide._CodePanes getCodePanes();


  /**
   * <p>
   * Getter method for the COM property "Windows"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide._Windows
   */

  @DISPID(110) //= 0x6e. The runtime will prefer the VTID if present
  @VTID(11)
  com.exceljava.com4j.vbide._Windows getWindows();


  /**
   * <p>
   * Getter method for the COM property "Events"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide.Events
   */

  @DISPID(111) //= 0x6f. The runtime will prefer the VTID if present
  @VTID(12)
  com.exceljava.com4j.vbide.Events getEvents();


  /**
   * <p>
   * Getter method for the COM property "ActiveVBProject"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide._VBProject
   */

  @DISPID(201) //= 0xc9. The runtime will prefer the VTID if present
  @VTID(13)
  com.exceljava.com4j.vbide._VBProject getActiveVBProject();


  /**
   * <p>
   * Setter method for the COM property "ActiveVBProject"
   * </p>
   * @param lppptReturn Mandatory com.exceljava.com4j.vbide._VBProject parameter.
   */

  @DISPID(201) //= 0xc9. The runtime will prefer the VTID if present
  @VTID(14)
  void setActiveVBProject(
    com.exceljava.com4j.vbide._VBProject lppptReturn);


  /**
   * <p>
   * Getter method for the COM property "SelectedVBComponent"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide._VBComponent
   */

  @DISPID(202) //= 0xca. The runtime will prefer the VTID if present
  @VTID(15)
  com.exceljava.com4j.vbide._VBComponent getSelectedVBComponent();


  /**
   * <p>
   * Getter method for the COM property "MainWindow"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide.Window
   */

  @DISPID(204) //= 0xcc. The runtime will prefer the VTID if present
  @VTID(16)
  com.exceljava.com4j.vbide.Window getMainWindow();


  /**
   * <p>
   * Getter method for the COM property "ActiveWindow"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide.Window
   */

  @DISPID(205) //= 0xcd. The runtime will prefer the VTID if present
  @VTID(17)
  com.exceljava.com4j.vbide.Window getActiveWindow();


  /**
   * <p>
   * Getter method for the COM property "ActiveCodePane"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide._CodePane
   */

  @DISPID(206) //= 0xce. The runtime will prefer the VTID if present
  @VTID(18)
  com.exceljava.com4j.vbide._CodePane getActiveCodePane();


  /**
   * <p>
   * Setter method for the COM property "ActiveCodePane"
   * </p>
   * @param ppCodePane Mandatory com.exceljava.com4j.vbide._CodePane parameter.
   */

  @DISPID(206) //= 0xce. The runtime will prefer the VTID if present
  @VTID(19)
  void setActiveCodePane(
    com.exceljava.com4j.vbide._CodePane ppCodePane);


  /**
   * <p>
   * Getter method for the COM property "Addins"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide._AddIns
   */

  @DISPID(209) //= 0xd1. The runtime will prefer the VTID if present
  @VTID(20)
  com.exceljava.com4j.vbide._AddIns getAddins();


  // Properties:
}
