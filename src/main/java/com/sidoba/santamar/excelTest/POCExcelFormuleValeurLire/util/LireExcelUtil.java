package com.sidoba.santamar.excelTest.POCExcelFormuleValeurLire.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.constant.ErrorConstant;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class LireExcelUtil {
    public static String creerTableValeurs(MultipartFile multipartFile) {
        try {
            File fichier = File.createTempFile("temp", multipartFile.getOriginalFilename());
            multipartFile.transferTo(fichier);

            String nomFichier = fichier.getName();
            String extensionFichier = "";

            int indexDernierPoint = nomFichier.lastIndexOf('.');
            if (indexDernierPoint > 0 && indexDernierPoint < nomFichier.length() - 1) {
                extensionFichier = nomFichier.substring(indexDernierPoint + 1).toLowerCase();
            }

            FileInputStream inputStream = new FileInputStream(fichier);

            Workbook workbook = null;
            if (extensionFichier.equals("xls")) {
                workbook = new HSSFWorkbook(inputStream);
            }
            else{
                workbook = new XSSFWorkbook(inputStream);
            }

            Sheet sheet = workbook.getSheetAt(0);

            int premierLigne = sheet.getFirstRowNum();
            int dernierLigne = sheet.getLastRowNum();

            StringBuilder tableHTML = new StringBuilder();
            tableHTML.append("<table>")
                    .append("<tr><th>Valeur</th><th>Type</th><th>Cellule Excel</th></tr>");

            //Nous passons en revue toutes les lignes
            for (int ligneIdx = premierLigne; ligneIdx <= dernierLigne; ligneIdx++) {
                Row row = sheet.getRow(ligneIdx);

                if (row == null) {
                    continue;
                }

                int premierCellule = row.getFirstCellNum();
                int dernierCellule = row.getLastCellNum();

                for(int celluleIndex = premierCellule; celluleIndex < dernierCellule; celluleIndex++) {
                    Cell cellule = row.getCell(celluleIndex);
                    if (cellule == null || cellule.getCellType() == CellType.BLANK) {
                        continue;
                    }

                    if (cellule.getCellType() != CellType.FORMULA) {
                        continue;
                    }

                    switch (cellule.getCachedFormulaResultType()) {
                        case BOOLEAN:
                            tableHTML.append("<tr><td>" + cellule.getBooleanCellValue() + "</td><td>Booléen</td><td>" +
                                    CellReference.convertNumToColString(cellule.getColumnIndex()) + (cellule.getRowIndex() + 1) + "</td></tr>");
                            break;
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(cellule)) {
                                tableHTML.append("<tr><td>" + cellule.getDateCellValue() + "</td><td>Date</td><td>" +
                                        CellReference.convertNumToColString(cellule.getColumnIndex()) + (cellule.getRowIndex() + 1) + "</td></tr>");
                            }
                            else {
                                tableHTML.append("<tr><td>" + cellule.getNumericCellValue() + "</td><td>Numérique</td><td>" +
                                        CellReference.convertNumToColString(cellule.getColumnIndex()) + (cellule.getRowIndex() + 1) + "</td></tr>");
                            }
                            break;
                        case STRING:
                            tableHTML.append("<tr><td>" + cellule.getRichStringCellValue() + "</td><td>String</td><td>" +
                                    CellReference.convertNumToColString(cellule.getColumnIndex()) + (cellule.getRowIndex() + 1) + "</td></tr>");
                            break;
                        case ERROR:
                            tableHTML.append("<tr style=\"font-weight: bold; color: red;\"><td>" + ErrorConstant.valueOf(cellule.getErrorCellValue()).getText() + "</td><td>ERREUR</td><td>" +
                                    CellReference.convertNumToColString(cellule.getColumnIndex()) + (cellule.getRowIndex() + 1) + "</td></tr>");
                            break;

                    }
                }

            }
            tableHTML.append("</table>");
            workbook.close();
            inputStream.close();

            return tableHTML.toString();

        } catch (IOException e) {
            e.printStackTrace();
        }

        return null;
    }
}
