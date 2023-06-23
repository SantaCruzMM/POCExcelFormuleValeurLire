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
                    .append("<tr><th>Formule</th><th>Valeur</th><th>Type</th><th>Cellule Excel</th></tr>");

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

                    String tableLigneOuverture = "<tr>";
                    String tableLigneFermeture = "</tr>";
                    String tableCelluleOuverture = "<td>";
                    String tableCelluleFermeture = "</td>";

                    StringBuilder stringBuilder = new StringBuilder();

                    switch (cellule.getCachedFormulaResultType()) {
                        case BOOLEAN:
                            stringBuilder.append(tableLigneOuverture)
                                    .append(tableCelluleOuverture)
                                    .append(cellule.getCellFormula())
                                    .append(tableCelluleFermeture)
                                    .append(tableCelluleOuverture)
                                    .append(cellule.getBooleanCellValue())
                                    .append(tableCelluleFermeture)
                                    .append(tableCelluleOuverture)
                                    .append("Booléen");
                            break;
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(cellule)) {
                                stringBuilder.append(tableLigneOuverture)
                                        .append(tableCelluleOuverture)
                                        .append(cellule.getCellFormula())
                                        .append(tableCelluleFermeture)
                                        .append(tableCelluleOuverture)
                                        .append(cellule.getDateCellValue())
                                        .append(tableCelluleFermeture)
                                        .append(tableCelluleOuverture)
                                        .append("Date");
                            }
                            else {
                                stringBuilder.append(tableLigneOuverture)
                                        .append(tableCelluleOuverture)
                                        .append(cellule.getCellFormula())
                                        .append(tableCelluleFermeture)
                                        .append(tableCelluleOuverture)
                                        .append(cellule.getNumericCellValue())
                                        .append(tableCelluleFermeture)
                                        .append(tableCelluleOuverture)
                                        .append("Numérique");
                            }
                            break;
                        case STRING:
                            stringBuilder.append(tableLigneOuverture)
                                    .append(tableCelluleOuverture)
                                    .append(cellule.getCellFormula())
                                    .append(tableCelluleFermeture)
                                    .append(tableCelluleOuverture)
                                    .append(cellule.getRichStringCellValue())
                                    .append(tableCelluleFermeture)
                                    .append(tableCelluleOuverture)
                                    .append("String");
                            break;
                        case ERROR:
                            stringBuilder.append("<tr style=\"font-weight: bold; color: red;\">")
                                    .append(tableCelluleOuverture)
                                    .append(cellule.getCellFormula())
                                    .append(tableCelluleFermeture)
                                    .append(tableCelluleOuverture)
                                    .append(ErrorConstant.valueOf(cellule.getErrorCellValue()).getText())
                                    .append(tableCelluleFermeture)
                                    .append(tableCelluleOuverture)
                                    .append("ERREUR");
                            break;

                    }
                    stringBuilder.append(tableCelluleFermeture)
                            .append(tableCelluleOuverture)
                            .append(CellReference.convertNumToColString(cellule.getColumnIndex()) + (cellule.getRowIndex() + 1))
                            .append(tableCelluleFermeture)
                            .append(tableLigneFermeture);

                    tableHTML.append(stringBuilder.toString());
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
