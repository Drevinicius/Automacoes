package com.automacao_excel.automacao;

import javax.swing.JOptionPane;

/**
 * 
 * @author anvnc
*/

public class Alert{
    void showError(String title, String txt){
        JOptionPane.showMessageDialog(null,
                    txt,
                    title,
                    JOptionPane.ERROR_MESSAGE);
    }

    void showInfo(String title, String txt){
        JOptionPane.showMessageDialog(null,
                    txt,
                    title,
                    JOptionPane.INFORMATION_MESSAGE);
    }
}
