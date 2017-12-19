package org.jodconverter;

import org.jodconverter.office.*;
import org.jodconverter.office.DefaultOfficeManager;
import org.jodconverter.office.OfficeException;
import org.jodconverter.office.OfficeManager;

import java.io.File;

public class Converter {
    // Create an office manager using the default configuration.
    // The default port is 2002.
    private static OfficeManager officeManager;

    static {
        try {
            officeManager = DefaultOfficeManager.install();
            officeManager.start();
        } catch (OfficeException e) {
            System.out.printf("There was a problem starting the Office manager: ");
            e.printStackTrace();
            System.exit(1);
        }

    }

    public static void convert(String sourceFile, String destFile) {
        File inputFile = new File(sourceFile);
        File outputFile = new File(destFile);

        try {
            org.jodconverter.JODConverter.convert(inputFile).to(outputFile).execute();
        } catch (OfficeException e) {
            System.out.printf("There was a problem converting the document: ", e.getMessage());
        }
    }

    public static void shutdown() {
        try {
            officeManager.stop();
        } catch (OfficeException e) {
            System.out.println("Office failed to shut down properly");
        }
    }

    public static void main(String[] args) {
        String sourceDoc = "src/integTest/resources/documents/test.doc";
        String destDoc = "src/integTest/resources/documents/test.pdf";
        System.out.println("Converting docs");
        Converter.convert(sourceDoc, destDoc);
        System.out.println("Done");
    }


}
