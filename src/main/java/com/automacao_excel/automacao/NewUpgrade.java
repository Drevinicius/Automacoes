/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.automacao_excel.automacao;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.awt.Desktop;

import java.util.ArrayList;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.automacao_excel.automacao.Alert;

public class NewUpgrade {

    public void openFileToChanger(DirectoryPaths pathPrefeitura, DirectoryPaths pathSistema){
        try(FileInputStream fileInISS = new FileInputStream(pathPrefeitura.getPathFile())){
            try (FileInputStream fileInSIS = new FileInputStream(pathSistema.getPathFile())){
                Workbook workbookPrefeitura = new XSSFWorkbook(fileInISS);
                Workbook workbookSistema = new XSSFWorkbook(fileInSIS);

                Sheet sheetPrefeitura;
                Sheet sheetSistema = workbookSistema.getSheetAt(0);
                Sheet newsSheets;

                if(workbookPrefeitura.getSheetIndex("Notas analisadas") < 0){
                    newsSheets = workbookPrefeitura.createSheet("Notas analisadas");
                    sheetPrefeitura = workbookPrefeitura.getSheetAt(0);
                }else{
                    sheetPrefeitura = workbookPrefeitura.getSheetAt(0);
                    workbookPrefeitura.removeSheetAt(workbookPrefeitura.getSheetIndex("Notas analisadas"));
                    newsSheets = workbookPrefeitura.createSheet("Notas analisadas");
                }

                List<List<Object>> dataISS = new ArrayList<>();
                List<List<Object>> dataSIS = new ArrayList<>();
                List<List<Object>> newsDatas = new ArrayList<>();
                List<Object> cabecalho = new ArrayList<>();
                cabecalho.add("Num. NF");
                cabecalho.add("Data de emissão");
                cabecalho.add("CNPJ");
                cabecalho.add("Valor total");
                cabecalho.add("Valor do ISS");

                dataISS = new ManipulateData().getAllData(sheetPrefeitura);
                dataSIS = new ManipulateData().getAllData(sheetSistema);
                newsDatas.add(cabecalho);

                boolean existe = false;

                int nfISS = 0;
                double priceISS = 0.0;

                int nfSIS = 1;
                double priceSIS = 1.0;

                for(List<Object> ISS: dataISS){
                    existe = false;
                    if(ISS.size() >= 2 && !ISS.contains(null)){
                        nfISS = (Integer) ISS.get(0);
                        priceISS = (Double) ISS.get(3);
                        for(List<Object> SIS: dataSIS){
                            if(SIS.size() >= 2 && !SIS.contains(null)){
                                nfSIS = (Integer) SIS.get(0);
                                priceSIS = (Double) SIS.get(1);
                                if(nfISS == nfSIS && priceISS == priceSIS){
                                    existe = true;
                                    break;
                                }
                            }
                        }
                        if(!existe){
                            newsDatas.add(ISS);
                        }
                    }
                }

                new ManipulateData().SaveSheet(newsSheets, newsDatas);

                try(FileOutputStream fileOut = new FileOutputStream(pathPrefeitura.getPathFile())){
                    workbookPrefeitura.write(fileOut);
                }catch(IOException e){
                    String txt = "Feche o arquivo" + pathPrefeitura.getNameFile();
                    String title = "Arquivo aberto em outro processo"; 
                    new Alert().showError(title, txt);
                }finally{
                    workbookPrefeitura.close();
                    workbookSistema.close();
                } // Tentativa se salvar o arquivo
            }catch (Exception e){
                new Alert().showError("Arquivo não encontrado", "Selecione um arquivo válido");
            } //  Validação de arquivo XLSX
        }catch(IOException e){
                new Alert().showError("Arquivo não encontrado", "Selecione um arquivo válido");            
        }finally {
            String txt = "Arquivo " + pathPrefeitura.getNameFile() + " salvo";
            String title = "Processo finalizado";
            new Alert().showInfo(title, txt);
            try{
                Desktop.getDesktop().open(pathPrefeitura.getPathFile());
            }catch(IOException e){

            }
        }
    }
}
