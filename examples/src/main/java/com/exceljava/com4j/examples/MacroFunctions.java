package com.exceljava.com4j.examples;

import com.exceljava.jinx.ExcelAddIn;
import com.exceljava.jinx.ExcelMacro;

import com.exceljava.com4j.JinxBridge;
import com.exceljava.com4j.excel.*;
import com.exceljava.jinx.ExcelReference;
import com.exceljava.jinx.IUnknown;

import javax.swing.*;
import java.awt.*;

/**
 * Example macros that use com4j to call back into Excel
 * when triggered from an Excel control.
 */
public class MacroFunctions {
    private final ExcelAddIn xl;

    /**
     * Called by Jinx when binding the instance methods to Excel macros.
     *
     * Using a non-default constructor taking an ExcelAddIn is necessary
     * so that we can get hold of the COM Excel Application wrapper later.
     *
     * @param xl ExcelAddIn instance provided by Jinx.
     */
    public MacroFunctions(ExcelAddIn xl) {
        this.xl = xl;
    }

    /**
     * Example macro function to be bound to a checkbox.
     *
     * It sets the value of a named range with the name of the
     * checkbox + "_OUTPUT"
     */
    @ExcelMacro("jinx.checkbox_example")
    public void checkboxExample() {
        _Application app = JinxBridge.getApplication(xl);

        // get the checkbox that called this macro
        _Worksheet sheet = app.getActiveSheet().queryInterface(_Worksheet.class);
        CheckBoxes checkboxes = sheet.checkBoxes().queryInterface(CheckBoxes.class);
        CheckBox checkbox = checkboxes.item(app.getCaller()).queryInterface(CheckBox.class);

        // Find the named range for this checkbox
        String name = checkbox.getName();
        Range range = sheet.getRange(name + "_OUTPUT");

        // Set the cell value based on the checkbox state
        Object value = checkbox.getValue();
        if (value instanceof Double && (Double)value != 0.0) {
            range.setValue("Checked!");
        } else {
            range.setValue("Click the checkbox");
        }
    }

    /**
     * Example macro function to be bound to a scrollbar.
     *
     * It sets the value of a named range with the name of the
     * scrollbar + "_OUTPUT" to the current scrollbar value.
     */
    @ExcelMacro("jinx.scrollbar_example")
    public void scrollbarExample() {
        _Application app = JinxBridge.getApplication(xl);

        // Get the scrollbar that called this macro
        _Worksheet sheet = app.getActiveSheet().queryInterface(_Worksheet.class);
        ScrollBars scrollbars = sheet.scrollBars().queryInterface(ScrollBars.class);
        ScrollBar scrollbar = scrollbars.item(app.getCaller()).queryInterface(ScrollBar.class);

        // Find the named range for this scrollbar
        String name = scrollbar.getName();
        Range range = sheet.getRange(name + "_OUTPUT");

        // Set the cell value from the scrollbar value
        range.setValue(scrollbar.getValue());
    }

    @ExcelMacro(
            value = "jinx.show_object_example",
            shortcut = "Ctrl+Shift+I"
    )
    public void showObject() throws HeadlessException {
        // Get the current selection
        _Application app = JinxBridge.getApplication(xl);
        Range selection = app.getSelection().queryInterface(Range.class);

        // Ensure the cell is calculated
        selection.setFormula(selection.getFormula());
        selection.calculate();

        // Get an ExcelReference corresponding to the selection
        IUnknown unk = JinxBridge.getIUnknown(selection);
        ExcelReference cell = xl.getReference(unk);

        // Find the cached object for this cell
        Object cachedObject = xl.getCachedObject(cell);

        // Popup a non-modal dialog with the string representation of the object
        String message = cachedObject != null ? cachedObject.toString() : "NULL";
        JOptionPane pane = new JOptionPane(message, JOptionPane.INFORMATION_MESSAGE);
        JDialog dialog = pane.createDialog("Java Object");
        dialog.setModal(false);
        dialog.setVisible(true);
    }
}
