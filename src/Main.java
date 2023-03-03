/* Josue Fadul Mejia - T00062598 - 1P */

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.filechooser.FileSystemView;

public class Main {


    public static void main(String[] args) throws FileNotFoundException {

        JOptionPane.showMessageDialog(null, "Usuario. Tenga en cuenta que debera tener cerrado el archivo Excel que desea procesar.", "CERRAR EXCEL A PROCESAR",JOptionPane.WARNING_MESSAGE);
        String excelFilePath;
        JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
        FileNameExtensionFilter filter = new FileNameExtensionFilter("ARCHIVOS EXCEL", "xlsx", "xls");
        jfc.setFileFilter(filter);

        int returnValue = jfc.showOpenDialog(null);
        if (returnValue == JFileChooser.APPROVE_OPTION) {
            File selectedFile = jfc.getSelectedFile();
            excelFilePath = selectedFile.getAbsolutePath();

            try {
                BufferedReader br = new BufferedReader(new FileReader(excelFilePath));



                try {
                    FileInputStream inputStream = new FileInputStream(excelFilePath);
                    Workbook workbook = new XSSFWorkbook(inputStream);
                    Sheet sheet = workbook.getSheetAt(0);

                    List cellData = new ArrayList<>();

                    Iterator rowIterator = sheet.rowIterator();

                    while (rowIterator.hasNext()) {
                        XSSFRow hssfRow = (XSSFRow) rowIterator.next();
                        Iterator iterator = hssfRow.cellIterator();
                        List cellTemp = new ArrayList();

                        while (iterator.hasNext()) {
                            XSSFCell hssfCell = (XSSFCell) iterator.next();
                            cellTemp.add(hssfCell);
                        }
                        cellData.add(cellTemp);
                    }

                    if (excelFilePath != null) {
                        JOptionPane.showMessageDialog(null," El Archivo leido con exito. Ruta de la hoja de calculo:"+excelFilePath, "EXITO",JOptionPane.INFORMATION_MESSAGE );
                        JOptionPane.showMessageDialog(null,"Usuario. En caso de encontrar una inconsistencia en los datos calculados, rectificar los datos digitados en las notas. " +
                                "RECUERDE, que excel trabaja con PUNTOS (.) para los decimales y las notas deben estár comprendidas entre 0 y 5", "TENER EN CUENTA",JOptionPane.WARNING_MESSAGE);
                    }


                    int rowCount = 0;


                    //Fila 23, creacion de celdas e insercion de formulas en excel que cambiarán dependiendo los valores de los datos
                    Row rowQuiz1 = sheet.createRow(rowCount + 22);
                    Cell cellQuiz = rowQuiz1.createCell(0);
                    cellQuiz.setCellValue("Quiz 1");
                    Cell cellminq = rowQuiz1.createCell(1);
                    cellminq.setCellFormula(String.format("IF(AND(AND(D2>=0),AND(D3>=0), AND(D4>=0), AND(D5>=0), AND(D6>=0), AND(D7>=0), AND(D8>=0), AND(D9>=0),AND(D10>=0),AND(D11>=0),AND(D12>=0),AND(D13>=0),AND(D14>=0),AND(D15>=0),AND(D16>=0),AND(D17>=0),AND(D18>=0),AND(D19>=0), AND(D20>=0)),MIN(D2:D20), \"Algunas notas estan fuera de los limites\")"));
                    Cell cellmaxq = rowQuiz1.createCell(2);
                    cellmaxq.setCellFormula(String.format("IF(AND(AND(D2<=5),AND(D3<=5), AND(D4<=5), AND(D5<=5), AND(D6<=5), AND(D7<=5), AND(D8<=5), AND(D9<=5),AND(D10<=5),AND(D11<=5),AND(D12<=5),AND(D13<=5),AND(D14<=5),AND(D15<=5),AND(D16<=5),AND(D17<=5),AND(D18<=5),AND(D19<=5), AND(D20<=5)),MAX(D2:D20), \"Algunas notas estan fuera de los limites\")"));
                    Cell cellDesv = rowQuiz1.createCell(5);
                    cellDesv.setCellValue("Desviacion Estandar");
                    Cell cellvDesv = rowQuiz1.createCell(6);
                    cellvDesv.setCellFormula("STDEV(B32:B50)");

                    //Fila 24, creacion de celdas e insercion de formulas en excel que cambiarán dependiendo los valores de los datos
                    Row rowQuiz2 = sheet.createRow(rowCount + 23);
                    Cell cell2 = rowQuiz2.createCell(0);
                    cell2.setCellValue("Quiz 2");
                    Cell cellmixq1 = rowQuiz2.createCell(1);
                    cellmixq1.setCellFormula(String.format("IF(AND(AND(E2>=0),AND(E3>=0), AND(E4>=0), AND(E5>=0), AND(E6>=0), AND(E7>=0), AND(E8>=0), AND(E9>=0),AND(E10>=0),AND(E11>=0),AND(E12>=0),AND(E13>=0),AND(E14>=0),AND(E15>=0),AND(E16>=0),AND(E17>=0),AND(E18>=0),AND(E19>=0), AND(E20>=0)),MIN(E2:E20), \"Algunas notas estan fuera de los limites\")"));
                    Cell cellmax1 = rowQuiz2.createCell(2);
                    cellmax1.setCellFormula(String.format("IF(AND(AND(E2<=5),AND(E3<=5), AND(E4<=5), AND(E5<=5), AND(E6<=5), AND(E7<=5), AND(E8<=5), AND(E9<=5),AND(E10<=5),AND(E11<=5),AND(E12<=5),AND(E13<=5),AND(E14<=5),AND(E15<=5),AND(E16<=5),AND(E17<=5),AND(E18<=5),AND(E19<=5), AND(E20<=5)),MAX(E2:E20), \"Algunas notas estan fuera de los limites\")"));
                    Cell cellCapro = rowQuiz2.createCell(5);
                    cellCapro.setCellValue("Cantidad de Aprobados");
                    Cell cellvCapro = rowQuiz2.createCell(6);
                    cellvCapro.setCellFormula("COUNTIFS(C32:C50,\"=Aprobado\")");

                    //Fila 25, creacion de celdas e insercion de formulas en excel que cambiarán dependiendo los valores de los datos
                    Row rowQuiz3 = sheet.createRow(rowCount + 24);
                    Cell cell3 = rowQuiz3.createCell(0);
                    cell3.setCellValue("Quiz 3");
                    Cell cellminq2 = rowQuiz3.createCell(1);
                    cellminq2.setCellFormula(String.format("IF(AND(AND(F2>=0),AND(F3>=0), AND(F4>=0), AND(F5>=0), AND(F6>=0), AND(F7>=0), AND(F8>=0), AND(F9>=0),AND(F10>=0),AND(F11>=0),AND(F12>=0),AND(F13>=0),AND(F14>=0),AND(F15>=0),AND(F16>=0),AND(F17>=0),AND(F18>=0),AND(F19>=0), AND(F20>=0)),MIN(F2:F20), \"Algunas notas estan fuera de los limites\")"));
                    Cell cellmaxq2 = rowQuiz3.createCell(2);
                    cellmaxq2.setCellFormula(String.format("IF(AND(AND(F2<=5),AND(F3<=5), AND(F4<=5), AND(F5<=5), AND(F6<=5), AND(F7<=5), AND(F8<=5), AND(F9<=5),AND(F10<=5),AND(F11<=5),AND(F12<=5),AND(F13<=5),AND(F14<=5),AND(F15<=5),AND(F16<=5),AND(F17<=5),AND(F18<=5),AND(F19<=5), AND(F20<=5)),MAX(F2:F20), \"Algunas notas estan fuera de los limites\")"));
                    Cell cellCdesa = rowQuiz3.createCell(5);
                    cellCdesa.setCellValue("Cantidad de Reprobado");
                    Cell cellvCdesa = rowQuiz3.createCell(6);
                    cellvCdesa.setCellFormula("COUNTIFS(C32:C50,\"=Reaprobado\")");


                    //Fila 26, creacion de celdas e insercion de formulas en excel que cambiarán dependiendo los valores de los datos
                    Row rowTaller1 = sheet.createRow(rowCount + 25);
                    Cell cell4 = rowTaller1.createCell(0);
                    cell4.setCellValue("Taller 1");
                    Cell cellmint = rowTaller1.createCell(1);
                    cellmint.setCellFormula(String.format("IF(AND(AND(G2>=0),AND(G3>=0), AND(G4>=0), AND(G5>=0), AND(G6>=0), AND(G7>=0), AND(G8>=0), AND(G9>=0),AND(G10>=0),AND(G11>=0),AND(G12>=0),AND(G13>=0),AND(G14>=0),AND(G15>=0),AND(G16>=0),AND(G17>=0),AND(G18>=0),AND(G19>=0), AND(G20>=0)),MIN(G2:G20), \"Algunas notas estan fuera de los limites\")"));
                    Cell cellmaxt = rowTaller1.createCell(2);
                    cellmaxt.setCellFormula(String.format("IF(AND(AND(G2<=5),AND(G3<=5), AND(G4<=5), AND(G5<=5), AND(G6<=5), AND(G7<=5), AND(G8<=5), AND(G9<=5),AND(G10<=5),AND(G11<=5),AND(G12<=5),AND(G13<=5),AND(G14<=5),AND(G15<=5),AND(G16<=5),AND(G17<=5),AND(G18<=5),AND(G19<=5), AND(G20<=5)),MAX(G2:G20), \"Algunas notas estan fuera de los limites\")"));


                    //Fila 27, creacion de celdas e insercion de formulas en excel que cambiarán dependiendo los valores de los datos
                    Row rowTaller2 = sheet.createRow(rowCount + 26);
                    Cell cell5 = rowTaller2.createCell(0);
                    cell5.setCellValue("Taller 2");
                    Cell cellmint1 = rowTaller2.createCell(1);
                    cellmint1.setCellFormula(String.format("IF(AND(AND(H2>=0),AND(H3>=0), AND(H4>=0), AND(H5>=0), AND(H6>=0), AND(H7>=0), AND(H8>=0), AND(H9>=0),AND(H10>=0),AND(H11>=0),AND(H12>=0),AND(H13>=0),AND(H14>=0),AND(H15>=0),AND(H16>=0),AND(H17>=0),AND(H18>=0),AND(H19>=0), AND(H20>=0)),MIN(H2:H20), \"Algunas notas estan fuera de los limites\")"));
                    Cell cellmaxt1 = rowTaller2.createCell(2);
                    cellmaxt1.setCellFormula(String.format("IF(AND(AND(H2<=5),AND(H3<=5), AND(H4<=5), AND(H5<=5), AND(H6<=5), AND(H7<=5), AND(H8<=5), AND(H9<=5),AND(H10<=5),AND(H11<=5),AND(H12<=5),AND(H13<=5),AND(H14<=5),AND(H15<=5),AND(H16<=5),AND(H17<=5),AND(H18<=5),AND(H19<=5), AND(H20<=5)),MAX(H2:H20), \"Algunas notas estan fuera de los limites\")"));


                    //Fila 28, creacion de celdas e insercion de formulas en excel que cambiarán dependiendo los valores de los datos
                    Row rowParcial = sheet.createRow(rowCount + 27);
                    Cell cell6 = rowParcial.createCell(0);
                    cell6.setCellValue("Parcial");
                    Cell cellminp = rowParcial.createCell(1);
                    cellminp.setCellFormula(String.format("IF(AND(AND(C2>=0),AND(C3>=0), AND(C4>=0), AND(C5>=0), AND(C6>=0), AND(C7>=0), AND(C8>=0), AND(C9>=0),AND(C10>=0),AND(C11>=0),AND(C12>=0),AND(C13>=0),AND(C14>=0),AND(C15>=0),AND(C16>=0),AND(C17>=0),AND(C18>=0),AND(C19>=0), AND(C20>=0)),MIN(C2:C20), \"Algunas notas estan fuera de los limites\")"));
                    Cell cellmaxp = rowParcial.createCell(2);
                    cellmaxp.setCellFormula(String.format("IF(AND(AND(C2<=5),AND(C3<=5), AND(C4<=5), AND(C5<=5), AND(C6<=5), AND(C7<=5), AND(C8<=5), AND(C9<=5),AND(C10<=5),AND(C11<=5),AND(C12<=5),AND(C13<=5),AND(C14<=5),AND(C15<=5),AND(C16<=5),AND(C17<=5),AND(C18<=5),AND(C19<=5), AND(C20<=5)),MAX(C2:C20), \"Algunas notas estan fuera de los limites\")"));


                    //Fila 22, creacion de celdas e insercion de formulas en excel que cambiarán dependiendo los valores de los datos
                    Row rowFila22 = sheet.createRow(rowCount + 21);
                    Cell cellminimo = rowFila22.createCell(1);
                    cellminimo.setCellValue("Min");
                    Cell cellmaxaximo = rowFila22.createCell(2);
                    cellmaxaximo.setCellValue("Max");
                    Cell cellProm = rowFila22.createCell(5);
                    cellProm.setCellValue("Promedio General");
                    Cell cellvProm = rowFila22.createCell(6);
                    cellvProm.setCellFormula("SUM(B32:B50)/COUNT(B32:B50)");

                    //Encabezado
                    Row encabezado = sheet.createRow(rowCount + 30);
                    Cell celda1 = encabezado.createCell(0);
                    celda1.setCellValue("Codigo");
                    Cell celda2 = encabezado.createCell(1);
                    Cell aprobado = encabezado.createCell(2);
                    Cell promp = encabezado.createCell(4);
                    aprobado.setCellValue("Calidad");
                    celda2.setCellValue("Nota");
                    promp.setCellValue("Estudiantes con promedio perfecto");



                    //Ciclo para crear celdas e introducir la formula excel que variará dependiendo el parametro
                    for(int i = 2; i <21; i++) {

                        Row rowTotal6 = sheet.createRow(rowCount + 29+i);
                        Cell cellTotal6 = rowTotal6.createCell(1);

                        cellTotal6.setCellFormula(String.format("IF(AND(AND(C%d >=0, C%d<=5), AND(D%d >=0, D%d<=5), AND(E%d >=0, E%d<=5), AND(F%d >=0, F%d<=5), AND(G%d >=0, G%d<=5), AND(H%d >=0, H%d<=5)),  (C%d*0.5)+((D%d+E%d+F%d)/3)*0.3+((G%d+H%d)/2)*0.2, \"Revisar datos digitados\")", i,i,i,i,i,i,i,i,i,i,i,i,i,i,i,i,i,i));
                        Cell AoR = rowTotal6.createCell(2);
                        Cell codigo = rowTotal6.createCell(0);
                        Cell Estup = rowTotal6.createCell(4);

                        codigo.setCellFormula(String.format("A%d", i));
                        AoR.setCellFormula(String.format("IF(OR(B%d<3, B%d=\"Revisar datos digitados\"),\"Reaprobado\", \"Aprobado\")", 30+i, 30+i));
                        //AoR.setCellFormula("Reprobado");
                        Estup.setCellFormula(String.format("IF(B%d=5,B%d,\" \")",30+i, i));
                    }

                    //Ajustar automaticamente la celda al contexto
                    sheet.autoSizeColumn(0);
                    sheet.autoSizeColumn(1);
                    sheet.autoSizeColumn(2);
                    sheet.autoSizeColumn(3);
                    sheet.autoSizeColumn(4);
                    sheet.autoSizeColumn(5);

                    //Ruta donde se almacenara el archivo escrito
                    FileOutputStream outputStream = new FileOutputStream(excelFilePath);
                    workbook.write(outputStream);
                    outputStream.close();

                } catch (IOException | EncryptedDocumentException ex) {
                    ex.printStackTrace();
                }

            } catch (IOException ex) {
                JOptionPane.showMessageDialog(null,"No se ha encontrado el archivo.", "ERROR", JOptionPane.ERROR_MESSAGE);
            }


        }

    }
}