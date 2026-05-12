/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.automacao_excel.automacao;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.io.File;

/**
 *
 * @author anvnc
 */

public class DirectoryPaths {
    private final JFileChooser root = new JFileChooser();
    private File pathFile;
    private final FileNameExtensionFilter config = new FileNameExtensionFilter("Planilhas (*.xls, *.xlsx)", "xls", "xlsx");
    
    // Construtor que inizializa minha Janela main como NULL
    public DirectoryPaths(){
        this.root.setCurrentDirectory(new File("C:\\"));
        this.root.addChoosableFileFilter(this.config);
        this.root.setPreferredSize(new java.awt.Dimension(800, 600));
        this.root.setSize(800, 600);        
    }

    // Metodo que chama uma tela JFileChooser para o usuário escolhar um file e retorna seu endereço em String
    public String setPathFile(String title){
        String path = null;
        root.setDialogTitle(title);
        if(root.showOpenDialog(null) == JFileChooser.APPROVE_OPTION){
            this.pathFile = this.root.getSelectedFile();
            path = this.pathFile.getAbsolutePath();
        }
        return path;
    }
    // Abre um JFileChooser para que o usuario possa escolher onde salvar um arquivo
    public String setPathDirectory() {
        String path = null;
        // Configuração de titulo
        root.setDialogTitle("Selecione a pasta para salvar");
        // DEFINE O NOME PADRÃO DO ARQUIVO 
        root.setSelectedFile(new File("novaPlanilha.xlsx"));
        // Abre a janela de salvar (this.rootMain)
        if (this.root.showSaveDialog(null) == JFileChooser.APPROVE_OPTION) {
            this.pathFile = this.root.getSelectedFile();
            path = this.pathFile.getAbsolutePath();

            // GARANTE QUE A EXTENSÃO .xlsx EXISTA
            if (!path.toLowerCase().endsWith(".xlsx")) {
                path += ".xlsx";
                this.pathFile = new File(path);
            }
        }
        return path;
    }
    
    public void redefinePathFile(){
        this.pathFile = null;
    }

    // metodos GETTERs
    public File getPathFile(){
        return this.pathFile;
    }
    public String getNameFile(){
        return this.pathFile.getName();
    }
}
