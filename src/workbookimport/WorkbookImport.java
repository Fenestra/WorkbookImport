/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package workbookimport;

/**
 *
 * @author lee
 */

import com.westat.ExcelToJson;

public class WorkbookImport {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws Exception {
      String sourceName = "sampleImportFormatting.xlsx";
      String destDir = "./";
      if (args.length > 0) {
         sourceName = args[0];
         if (args.length > 1) {
            destDir = args[1];
         }
      }
      
      System.out.println("Executing conversion of Excel to Json...");
      System.out.println("syntax is: WorkbookImport sourceFilename destinationDirectory");
      System.out.println("source:"+sourceName);
      System.out.println("destDir:"+destDir);
      
      ExcelToJson conv = new ExcelToJson();
      conv.convertExcelToJson(sourceName, destDir);
 
    }
}
