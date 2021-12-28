package com.nickjmoss;

import java.awt.FlowLayout;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextArea;
import javax.swing.UIManager;
import javax.swing.WindowConstants;
import javax.swing.filechooser.FileNameExtensionFilter;

public class Frame {
    JFrame f;
    JPanel mainPanel;
    JLabel text;
    JTextArea textArea;
    JButton fileChooser;
    JButton decompress;
    JButton compress;

    public Frame() {
        f = new JFrame("Spreadsheet Compressor & Decompressor");
        f.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        mainPanel = new JPanel();

        mainPanel.setLayout(new GridLayout(4, 1));
        mainPanel.setBorder(BorderFactory.createEmptyBorder(15, 20, 20, 15));

        // input panel
        JPanel input = new JPanel();
        input.setLayout(new FlowLayout());
        // function buttons panel
        JPanel functions = new JPanel();
        functions.setLayout(new FlowLayout());
        // output panel
        JPanel output = new JPanel();
        functions.setLayout(new FlowLayout());

        // TextField for file data path
        text = new JLabel("file path...");

        // TextArea for output
        JLabel outputLabel = new JLabel("Output:");
        textArea = new JTextArea(30, 40);
        textArea.setWrapStyleWord(true);
        textArea.setLineWrap(true);
        textArea.setEditable(false);
        textArea.setBackground(UIManager.getColor("text.background"));

        // creating buttons
        fileChooser = new JButton("Choose a File...");
        decompress = new JButton("Decompress");
        compress = new JButton("Compress");

        // add buttons to the panel
        input.add(fileChooser);
        input.add(text);
        functions.add(decompress);
        functions.add(compress);
        output.add(textArea);
        mainPanel.add(input);
        mainPanel.add(functions);
        mainPanel.add(outputLabel);
        mainPanel.add(output);

        f.add(mainPanel);

        f.setSize(600, 400);
        f.setLocationRelativeTo(null);
        f.setVisible(true);
    }

    public void buttonActions() {
        ActionListener fileListener = new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                JFileChooser chooser = new JFileChooser();
                chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
                chooser.setCurrentDirectory(new File("/Users/nickmoss1999/projects/java_project/ops-compressor"));
                chooser.setFileFilter(new FileNameExtensionFilter("Excel files", "xlsx"));
                int result = chooser.showOpenDialog(null);

                if (result == JFileChooser.APPROVE_OPTION) {
                    text.setText(chooser.getSelectedFile().getAbsolutePath());
                }
            }
        };

        ActionListener decompressListener = new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                ExcelEditor editor = new ExcelEditor();
                String result = editor.decompress(text.getText());
                textArea.setText(result);
            }
        };

        ActionListener compressListener = new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                ExcelEditor editor = new ExcelEditor();
                String result = editor.compress(text.getText());
                textArea.setText(result);
            }
        };

        decompress.addActionListener(decompressListener);
        compress.addActionListener(compressListener);
        fileChooser.addActionListener(fileListener);
    }
};