package almacen;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Iterator;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Almacen {
  public static void main(String[] args) {
      
    StringBuilder sb = new StringBuilder();
    DataFormatter formatter = new DataFormatter();
    int cont = 0;
    boolean cumple_sap = false;
    boolean sap_especial=false;
    String CMAC = null;
    String MTA_MAC = null;
    String codigo_sap = null;
    String serie_equipo = null;
    String UA_DECO = null;
    String codigo_mac= null;
    
    
    try {
      String excelFilePath = "D:\\prueba.xlsx";
      FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
      XSSFWorkbook xSSFWorkbook = new XSSFWorkbook(inputStream);
      Sheet firstSheet = xSSFWorkbook.getSheetAt(0);
      Iterator<Row> iterator = firstSheet.iterator();
      
      Row nextRow = iterator.next();
      while (iterator.hasNext()) {
        nextRow = iterator.next();
        Iterator<Cell> cellIterator = nextRow.cellIterator();
        while (cellIterator.hasNext()) {
          Cell cell = cellIterator.next();
          
     //desde la columna cero hasta la columna 3 no se toma en cuenta   
          if (cell.getColumnIndex() < 4)
            continue; 
          
      //la columna 4 se valida el codigo SAP y se deja la variable codigo sap lista para usar
          if (cell.getColumnIndex() == 4) {
            codigo_sap = formatter.formatCellValue(cell);
            //si el codigo sap cumple con ciertos codigos de numeros se pone la variable contador a 1 ...
            if (codigo_sap.equals("1047000") || codigo_sap.equals("4047000") || codigo_sap.equals("4048632") || codigo_sap.equals("4053488"))
            
            sap_especial=  true;
            
            if (codigo_sap.endsWith("N") || codigo_sap.endsWith("N2")) {
              if (codigo_sap.endsWith("N")) {
                codigo_sap = "00000000000" + codigo_sap.substring(0, codigo_sap.length() - 1);
                continue;
              } 
              if (codigo_sap.endsWith("N2")) {
                codigo_sap = "00000000000" + codigo_sap.substring(0, codigo_sap.length() - 2);
                continue;
              } 
            } 
            codigo_sap = "00000000000" + codigo_sap;
            continue;
          } 
                      
   //se valida la columna 5 ( INVSN )  DONDE ESTA EL IPTV !!!!!!!!!!!!!
          if (cell.getColumnIndex() == 5) {
            String str = formatter.formatCellValue(cell);
            //si no tiene letras o caracter raro
            if (!str.matches("[a-zAZ0-Z0-9]*")) {
              sap_especial=false;
              break;
            } 
            
            if (str.startsWith("21") && str.length() > 18) {
              serie_equipo = str.substring(2, str.length());
              continue;
            } 
            serie_equipo = formatter.formatCellValue(cell);
            continue;
          } 
          
   // se valida columna 6 (XI_MTA_MAC_CM) ( si es iptv la columna anterior y esta esta vacia lo copiamos de la columna 8 creo)
          if (cell.getColumnIndex() == 6) {
            String str = formatter.formatCellValue(cell);
            str= str.replace("Ñ", "");
            str= str.replace(":","");
            if (!str.matches("[a-zA-Z0-9]*")) {
              sap_especial=false;
              break;
            }     
            CMAC = str;            
            continue;
          } 
          
    //se valida la columna 7  (XI_MTA_MAC )    
          if (cell.getColumnIndex() == 7) {
            String str = formatter.formatCellValue(cell);
            if (!str.matches("[a-zA-Z0-9]*")) {
              sap_especial=false;
              break;
            } 
            MTA_MAC = formatter.formatCellValue(cell);
            continue;
          } 
          
  //se valida la columna 8  ( XI_UNIT_ADDR )  tenia valores con ñ que deben editarse
  
          if (cell.getColumnIndex() == 8) {
            String str = formatter.formatCellValue(cell);
           /* if (!str.matches("[a-zA-Z0-9]*")) {
              sap_especial=false;
              break;
            } */
            UA_DECO = formatter.formatCellValue(cell);
            if (CMAC.equals("") && MTA_MAC.equals("") && UA_DECO.equals("")) {
              sap_especial = false;
              break;
            } 
            if (sap_especial == true) {
              if (CMAC.equals("")) {
                sb.append(codigo_sap + "\t" + serie_equipo + "\t" + MTA_MAC + "\t");
                sb.append(MTA_MAC + "\t" + UA_DECO + "\t");
                continue;
              } 
              sb.append(codigo_sap + "\t" + serie_equipo + "\t" + CMAC + "\t");
              sb.append(CMAC + "\t" + UA_DECO + "\t");
              continue;
            } 
            sb.append(codigo_sap + "\t" + serie_equipo + "\t" + CMAC + "\t");
            sb.append(MTA_MAC + "\t" + UA_DECO + "\t");
            continue;
          } 
          
        
   //no se toma en cuenta desde la columna 9 hasta la columna 12; se ignoran!!!!!
          if (cell.getColumnIndex() > 8 && cell.getColumnIndex() < 13)
            continue; 
          
   //columna numero 13 se imprime normal,se reinicia el sap_especial y se crea una nueva fila
          if (cell.getColumnIndex() == 13) {
            sb.append(formatter.formatCellValue(cell) + "\t");
            //restablecemos el sap_especial a su valor por defecto en falso
            sap_especial= false;
            //un salto de linea a para la siguiente fila
            sb.append(System.getProperty("line.separator"));
          } 
        } 
      } 
      
     
      
//aca se pasa al bloc  de notas y se cierra el writer ...  
      File file = new File("D:\\carga.txt");
      try (BufferedWriter writer = new BufferedWriter(new FileWriter(file))) {
        writer.write(sb.toString());
        JOptionPane.showMessageDialog(null, "Revisar carga.txt en disco D");
      } 
      try {
        FileOutputStream out = new FileOutputStream("D:\\prueba.xlsx");
        xSSFWorkbook.write(out);
        out.close();
      } catch (FileNotFoundException e) {
        JOptionPane.showMessageDialog(null, e.getMessage());
        e.printStackTrace();
      } catch (IOException e) {
        JOptionPane.showMessageDialog(null, e.getMessage());
        e.printStackTrace();
      } 
    } catch (Exception e) {
      JOptionPane.showMessageDialog(null, e.getMessage());
      e.printStackTrace();
    } 
  }
}