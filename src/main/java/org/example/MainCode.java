package org.example;

import javax.swing.*;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseEvent;
import java.awt.event.MouseAdapter;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MainCode extends JFrame implements ActionListener {
    JLabel settings, title, minimize, exit, fileName, fileID, department, location;
    JPanel titleBar, mapPanel;
    JTextField find;
    JButton searchBt, clear;

    MainCode() {
        setSize(1100, 900);
        setLocationRelativeTo(null);
        setUndecorated(true); // Remove title bar for custom styling
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setBackground(new Color(131, 131, 131, 100));
        setLayout(null);

        settings = new JLabel("≡");
        settings.setFont(new Font("Arial", Font.PLAIN, 20));
        settings.setForeground(new Color(166, 179, 179));
        settings.setBounds(20, 5, 20, 30);
        add(settings);

        title = new JLabel("Find My Doc's 5s");
        title.setFont(new Font("Arial", Font.PLAIN, 16));
        title.setForeground(new Color(166, 179, 179));
        title.setBounds(50, 5, 200, 30);
        add(title);


        Cursor hand = new Cursor(Cursor.HAND_CURSOR);

        minimize = new JLabel("–");
        minimize.setFont(new Font("Arial", Font.PLAIN,16));
        minimize.setBounds(1040, 5, 16, 30);
        minimize.setForeground(Color.WHITE);
        minimize.setCursor(hand);
        add(minimize);

        minimize.addMouseListener(new MouseAdapter () {
            @Override
            public void mouseClicked(MouseEvent e) {
                setState(JFrame.ICONIFIED);
            }

            @Override
            public void mouseEntered(MouseEvent e) {
                minimize.setForeground(Color.GRAY);
            }

            @Override
            public void mouseExited(MouseEvent e) {
                minimize.setForeground(Color.WHITE);
            }
        });

        exit = new JLabel("X");
        exit.setFont(new Font("Arial", Font.PLAIN,16));
        exit.setBounds(1065, 5, 16, 30);
        exit.setForeground(Color.WHITE);
        exit.setCursor(hand);
        add(exit);

        exit.addMouseListener(new MouseAdapter () {
            @Override
            public void mouseClicked(MouseEvent e) {
                System.exit(0);
            }

            @Override
            public void mouseEntered(MouseEvent e) {
                exit.setForeground(Color.RED);
            }

            @Override
            public void mouseExited(MouseEvent e) {
                exit.setForeground(Color.WHITE);
            }
        });


        titleBar = new JPanel();
        titleBar.setBounds(0, 0, 1100, 40);
        titleBar.setBackground(new Color(45, 45, 48, 230));
        add(titleBar);

        ImageIcon rightImg = new ImageIcon(ClassLoader.getSystemResource("imgs/officeFiles.jpg"));
        Image i3 = rightImg.getImage().getScaledInstance(800, 800, Image.SCALE_SMOOTH);
        ImageIcon ic3 = new ImageIcon(i3);
        JLabel label3 = new JLabel(ic3);
        label3.setBounds(-20, 20, 900, 900);
        add(label3);

        mapPanel = new JPanel();
        mapPanel.setBounds(10, 50, 1080, 840);
        mapPanel.setBackground(new Color(45,45,45));
        mapPanel.setLayout(null);
        add(mapPanel);

        find = new JTextField();
        find.setFont(new Font("Times new Roman", Font.PLAIN, 16));
        find.setBounds(830, 220, 240, 30);
        find.setBackground(new Color(64, 64, 67));
        find.setForeground(new Color(201, 213,  213));
        mapPanel.add(find);

        searchBt = new JButton("Search");
        searchBt.setFont(new Font("Times new roman", Font.BOLD, 18));
        searchBt.setBackground(new Color(146, 158, 158));
        searchBt.setForeground(new Color(45,45,45));
        searchBt.setBounds(850, 260, 200, 30);
        searchBt.setCursor(hand);
        mapPanel.add(searchBt);

        searchBt.addMouseListener(new MouseAdapter () {
            @Override
            public void mouseEntered(MouseEvent e) {
                searchBt.setBackground(new Color(121, 221, 136));
            }

            @Override
            public void mouseExited(MouseEvent e) {
                searchBt.setBackground(new Color(146, 158, 158));
            }
        });

        clear = new JButton("Clear");
        clear.setFont(new Font("Times new roman", Font.BOLD, 18));
        clear.setBackground(new Color(158, 146, 146));
        clear.setForeground(new Color(45,45,45));
        clear.setBounds(850, 300, 200, 30);
        clear.setCursor(hand);
        mapPanel.add(clear);
        clear.addActionListener(this);

        clear.addMouseListener(new MouseAdapter () {
            @Override
            public void mouseEntered(MouseEvent e) {
                clear.setBackground(new Color(213, 135, 135));
            }

            @Override
            public void mouseExited(MouseEvent e) {
                clear.setBackground(new Color(158, 146, 146));
            }
        });

        fileName = new JLabel(" File Name : ");
        fileName.setFont(new Font("Arial", Font.PLAIN, 16));
        fileName.setForeground(new Color(166, 179, 179));
        fileName.setBounds(837, 40, 200, 30);
        mapPanel.add(fileName);

        fileID = new JLabel("ID :");
        fileID.setFont(new Font("Arial", Font.PLAIN, 16));
        fileID.setForeground(new Color(166, 179, 179));
        fileID.setBounds(897, 70, 200, 30);
        mapPanel.add(fileID);

        department = new JLabel("Department :");
        department.setFont(new Font("Arial", Font.PLAIN, 16));
        department.setForeground(new Color(166, 179, 179));
        department.setBounds(830, 100, 200, 30);
        mapPanel.add(department);

        location = new JLabel("Location :");
        location.setFont(new Font("Arial", Font.PLAIN, 16));
        location.setForeground(new Color(166, 179, 179));
        location.setBounds(830, 130, 200, 30);
        mapPanel.add(location);


        setVisible(true);
    }

    public void actionPerformed(ActionEvent e){
        if (e.getSource() == clear) {
            find.setText("");
            fileName.setText(" File Name : ");
            fileID.setText("ID : ");
            department.setText("Department : ");
            location.setText("Location : ");
        } else if (e.getSource() == searchBt) {
            String searchText = find.getText();
            if (!searchText.isEmpty()) {
                readExcelAndDisplay(searchText);
            }
        }
    }

    private void readExcelAndDisplay(String searchText) {
        String filePath = "/excelFile/data.xlsx";  // Path inside resources folder

        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                Cell cell = row.getCell(0);
                if (cell != null && cell.toString().equalsIgnoreCase(searchText)) {
                    fileName.setText("Material: " + cell.toString());
                    fileID.setText("ID: " + row.getCell(1));
                    department.setText("Department: " + row.getCell(2));
                    location.setText("Location: " + row.getCell(3));
                    return;
                }
            }
            JOptionPane.showMessageDialog(this, "Material not found in Excel.", "Search Result", JOptionPane.INFORMATION_MESSAGE);
        } catch (IOException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(this, "Error reading Excel file.", "Error", JOptionPane.ERROR_MESSAGE);
        }
    }


    public static void main(String[] args) {
        new MainCode();
    }
}