/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.automacao_excel.automacao;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import java.util.List;
import java.util.ArrayList;
import java.util.Date;
import java.text.ParseException;
import java.text.SimpleDateFormat;

/**
 *
 * @author anvnc
 */

public class ManipulateData {
    public Date converterParaDate(String valor){
        // Lista de formatos comuns (adicione conforme necessário)
        String[] formatos = {
                "dd/MM/yyyy",
                "dd-MM-yyyy",
                "MM/dd/yyyy",
                "dd/MM/yy",
                "yyyy/MM/dd"
        };

        for (String formato : formatos) {
            try {
                SimpleDateFormat sdf = new SimpleDateFormat(formato);
                sdf.setLenient(false); // Validação rigorosa
                return sdf.parse(valor);
            } catch (ParseException e) {
                continue;
            }
        }
        throw new IllegalArgumentException("Erro, man");
    }

    public List<List<Object>> getAllData(Sheet sheet){
        List<List<Object>> allData = new ArrayList<>();
        double priceNF = 0.0;
        int numberNF = 0;
        String CNPJ = "";
        double iss = 0.0;
        String dateNF = "new Date()";

        for(Row rows: sheet){
            List<Object> newDate = new ArrayList<>();
            for(Cell cells: rows){
                if(rows.getPhysicalNumberOfCells() >= 42){
                    switch (cells.getColumnIndex()){
                        case 3:
                            try{
                                numberNF = Integer.parseInt(cells.getStringCellValue().trim());
                                newDate.add(numberNF);
                            }catch (Exception e){
                                newDate.add(null);
                            }
                            break;
                        case 16:
                            switch (cells.getCellType()){
                                case NUMERIC:
                                    priceNF = cells.getNumericCellValue();
                                    newDate.add(priceNF);
                                    break;
                            }
                            break;
                    }
                }
                else if(rows.getPhysicalNumberOfCells() >= 12){
                    switch (cells.getColumnIndex()){
                        case 1:
                            try{
                                numberNF = Integer.parseInt(cells.getStringCellValue().trim());
                                newDate.add(numberNF);
                            }catch(Exception e){
                                newDate.add(null);
                            }
                            break;
                        case 3:
                                try{
                                    dateNF = cells.getStringCellValue();
                                    newDate.add(dateNF);
                                }catch(Exception e){
                                    newDate.add(null);
                                }
                            break;
                        case 5:
                            CNPJ = cells.getStringCellValue();
                            newDate.add(CNPJ);
                           break;
                        case 7:
                            try{
                                priceNF = cells.getNumericCellValue();
                                newDate.add(priceNF);
                            }catch(Exception e){
                                newDate.add(null);
                            }
                            break;
                        case 9:
                            switch (cells.getCellType()){
                                case NUMERIC:
                                    iss = cells.getNumericCellValue();
                                    newDate.add(iss);
                                    break;
                                default:
                                    newDate.add(null);
                                    break;
                            }
                            break;
                    }
                }
                if(!allData.contains(newDate) && !newDate.contains(null)){
                    allData.add(newDate);
                }
            }

        }
        return allData;
    }

    public void SaveSheet(Sheet sheet, List<List<Object>> dataToSave){
        int interador = 0;

        for(List<Object> lista: dataToSave){
            Row novaLinha = sheet.createRow(interador++);
            for(int i = 0; i < lista.size(); i++){
                Cell cell = novaLinha.createCell(i);
                if(i == 3 || i == 4){
                    String novo = (String)lista.get(i).toString();
                    novo = novo.replace(".", ",");
                    cell.setCellValue(novo);
                    continue;
                }
                cell.setCellValue((String)lista.get(i).toString());
            }
        }
    }
}
