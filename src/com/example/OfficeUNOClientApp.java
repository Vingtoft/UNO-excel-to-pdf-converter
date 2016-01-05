package com.example;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.container.XNameAccess;
import com.sun.star.container.XNameContainer;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XController;
import com.sun.star.frame.XModel;
import com.sun.star.frame.XStorable;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.sheet.XSpreadsheetView;
import com.sun.star.style.HorizontalAlignment;
import com.sun.star.style.XStyle;
import com.sun.star.style.XStyleFamiliesSupplier;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.view.PaperOrientation;
import com.sun.star.view.XPrintable;
import java.io.File;
import java.util.Scanner;

/**
 *
 * @author vingtoft
 */
public class OfficeUNOClientApp {

    /**
     * Creates a new instance of OfficeUNOClientApp
     */
    public OfficeUNOClientApp() {
    }

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        File[] files = new File("/Users/vingtoft/Documents/Development/Vingtoft Development/TopTrack/reports/").listFiles();
        for (File file : files) {
            String name = file.getName();
            String x;
            try {
                x = name.substring(name.lastIndexOf(".") + 1);
            } catch (Exception e) {
                x = "";
            }
            if (x.equals("xlsx") && !file.isDirectory()) {
                System.out.println(file.getAbsolutePath());
                String pathToFile = file.getAbsolutePath();

                String loadUrl = "file://" + pathToFile;
                String storeUrl = "file:///Users/vingtoft/Documents/Development/Vingtoft Development/TopTrack/reports/pdf/" + name.replace("xlsx", "pdf");

                System.out.println(storeUrl);

                try {
                    XComponentContext xContext = Bootstrap.bootstrap();
                    XMultiComponentFactory xMultiComponentFactory = xContext.getServiceManager();
                    XComponentLoader xcomponentloader = (XComponentLoader) UnoRuntime.queryInterface(XComponentLoader.class, xMultiComponentFactory.createInstanceWithContext("com.sun.star.frame.Desktop", xContext));

                    PropertyValue[] a = new PropertyValue[1];
                    a[0] = new PropertyValue();
                    a[0].Name = "Hidden";
                    a[0].Value = true;

                    Object objectDocumentToStore = xcomponentloader.loadComponentFromURL(loadUrl, "_blank", 0, a);

                    // Create PDF filter data
                    PropertyValue pdfFilterData[] = new PropertyValue[20];

                    PropertyValue[] conversionProperties = new PropertyValue[3];
                    conversionProperties[0] = new PropertyValue();
                    conversionProperties[0].Name = "FilterName";
                    conversionProperties[0].Value = "calc_pdf_Export";

                    XStyleFamiliesSupplier xSupplier = (XStyleFamiliesSupplier) UnoRuntime.queryInterface(XStyleFamiliesSupplier.class, objectDocumentToStore);
                    // Use the StyleFamiliesSupplier interface to get the XNameAccess interface of the
                    // actual style families
                    XNameAccess xFamilies = (XNameAccess) UnoRuntime.queryInterface(XNameAccess.class, xSupplier.getStyleFamilies());
                    // Access the 'PageStyles' Family
                    XNameContainer xFamily = (XNameContainer) UnoRuntime.queryInterface(XNameContainer.class, xFamilies.getByName("PageStyles"));
                    XSpreadsheetDocument spreadsheetDocument = (XSpreadsheetDocument) UnoRuntime.queryInterface(XSpreadsheetDocument.class, objectDocumentToStore);
                    XModel model = (XModel) UnoRuntime.queryInterface(XModel.class, spreadsheetDocument);
                    XController controller = model.getCurrentController();
                    XSpreadsheetView view = (XSpreadsheetView) UnoRuntime.queryInterface(XSpreadsheetView.class,
                            controller);
                    XPropertySet xps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, view.getActiveSheet());
                    
                    
                    
                    String pageStyleName = xps.getPropertyValue("PageStyle") + "";

                    XStyle xStyle = (XStyle) UnoRuntime.queryInterface(XStyle.class, xFamily.getByName(pageStyleName));
                    XPropertySet xStyleProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xStyle);
                    xStyleProps.setPropertyValue("ScaleToPages", new Short((short) 1));
                    xStyleProps.setPropertyValue("LeftMargin", new Short((short) 300));
                    xStyleProps.setPropertyValue("RightMargin", new Short((short) 300));
                    xStyleProps.setPropertyValue("BottomMargin", new Short((short) 300));
                    xStyleProps.setPropertyValue("TopMargin", new Short((short) 300));
                    
                    XPrintable xPrintable = (XPrintable) UnoRuntime.queryInterface(XPrintable.class, objectDocumentToStore);
                    PropertyValue[] printerDesc = new PropertyValue[1];
                    // Paper orientation
                    printerDesc[0] = new PropertyValue();
                    printerDesc[0].Name = "PaperOrientation";
                    printerDesc[0].Value = PaperOrientation.LANDSCAPE;
                    xPrintable.setPrinter(printerDesc);
                    
                    XStorable xstorable = (XStorable) UnoRuntime.queryInterface(XStorable.class, objectDocumentToStore);
                    xstorable.storeToURL(storeUrl, conversionProperties);

                } catch (java.lang.Exception e) {
                    e.printStackTrace();
                    System.exit(0);
                } finally {

                }
            }
        }
        
        System.exit(0);
    }
}
