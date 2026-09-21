package com.exceljava.com4j.examples;

import com.exceljava.jinx.ExcelAddIn;
import com.exceljava.jinx.ExcelMenu;

import com.exceljava.com4j.JinxBridge;
import com.exceljava.com4j.excel.*;
import com4j.COM4J;
import com4j.Com4jObject;
import com4j.util.ComObjectCollector;

/**
 * Example menu functions that use com4j to call back into Excel
 * when triggered from an Excel menu.
 */
public class MenuFunctions {
    private final ExcelAddIn xl;

    /**
     * Called by Jinx when binding the instance methods to Excel menus.
     *
     * Using a non-default constructor taking an ExcelAddIn is necessary
     * so that we can get hold of the COM Excel Application wrapper later.
     *
     * @param xl ExcelAddIn instance provided by Jinx.
     */
    public MenuFunctions(ExcelAddIn xl) {
        this.xl = xl;
    }

    /**
     * Example menu function that takes the current selection and sets
     * the background colors randomly.
     *
     * This will appear as a menu item under 'Jinx' in the Add-Ins tab
     * in Excel.
     */
    @ExcelMenu(
        value = "Randomise Colors",
        subMenu = "Com4J Examples"
    )
    public void randomiseColors() {
        // Collect the COM objects created by this method so that they can
        // all be disposed of before returning
        ComObjectCollector collector = new ComObjectCollector();
        COM4J.addListener(collector);

        try {
            _Application app = JinxBridge.getApplication(xl);

            Range active = app.getSelection().queryInterface(Range.class);
            for (int row = 1; row <= active.getRows().getCount(); row++) {
                for (int col = 1; col <= active.getColumns().getCount(); col++) {
                    Range cell = ((Com4jObject)active.getItem(row, col)).queryInterface(Range.class);

                    int red = (int)(Math.random() * 255);
                    int green = (int)(Math.random() * 255);
                    int blue = (int)(Math.random() * 255);

                    int color = red | (green << 8) | (blue << 16);
                    cell.getInterior().setColor(color);
                }
            }
        }
        finally {
            collector.disposeAll();
            COM4J.removeListener(collector);
        }
    }
}
