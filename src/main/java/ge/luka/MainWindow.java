package ge.luka;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.List;

public class MainWindow extends JFrame implements ActionListener {

    private JTextField sourceTextField = new JTextField(30);
    private JTextField destinationTextField = new JTextField(25);

    private JButton chooseSourceButton = new JButton("Choose source file");
    private JButton chooseDestinationButton = new JButton("choose destination folder");

    private JButton generateButton = new JButton("Generate");

    private JPanel errorMessagesPanel = new JPanel();

    public MainWindow() {
        super("Document generator");
        setSize(1200, 800);
        setResizable(true);

        setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        setLayout(new BorderLayout());

        JPanel mainPanel = new JPanel();
        add(mainPanel);
        mainPanel.setLayout(new GridLayout(5, 1));

        addEmptyPanel(mainPanel);
        addSourceLayer(mainPanel);
        addDestinationLayer(mainPanel);
        addGenerateButton(mainPanel);
        addErrorMessagesPanel(mainPanel);

        setLocationRelativeTo(null);
        setVisible(true);
    }

    private void addErrorMessagesPanel(JPanel mainPanel) {
        errorMessagesPanel.setLayout(new GridLayout(0, 1));
        JScrollPane jScrollPane = new JScrollPane(errorMessagesPanel,
                JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED,
                JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
        mainPanel.add(jScrollPane);
    }

    private void addEmptyPanel(JPanel mainPanel) {
        JPanel panel = new JPanel();
        mainPanel.add(panel);
    }

    private void addGenerateButton(JPanel mainPanel) {
        JPanel currentPanel = new JPanel();
        mainPanel.add(currentPanel);

        currentPanel.add(generateButton);
        generateButton.addActionListener(this);
    }

    private void addSourceLayer(JPanel mainPanel) {
        JPanel currentPanel = new JPanel();
        mainPanel.add(currentPanel);

        currentPanel.add(chooseSourceButton);
        chooseSourceButton.addActionListener(this);
        currentPanel.add(sourceTextField);
    }

    private void addDestinationLayer(JPanel mainPanel) {
        JPanel currentPanel = new JPanel();
        mainPanel.add(currentPanel);

        currentPanel.add(chooseDestinationButton);
        chooseDestinationButton.addActionListener(this);
        currentPanel.add(destinationTextField);
    }

    private void drawErrorMessages() {
        errorMessagesPanel.removeAll();
        String[] splits = ExcelReader.errorMessages.split("\n");
        if (splits.length == 0) {
            return;
        }
        for (int i = 0; i < splits.length; i++) {
            errorMessagesPanel.add(new JLabel(splits[i]));
        }
        revalidate();
        repaint();
    }

    @Override
    public void actionPerformed(ActionEvent e) {
        if (e.getSource() == chooseSourceButton) {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
            fileChooser.setFileFilter(new FileNameExtensionFilter(null, "xlsx"));
            int option = fileChooser.showOpenDialog(this);
            if (option == JFileChooser.APPROVE_OPTION) {
                File file = fileChooser.getSelectedFile();
                sourceTextField.setText(file.getAbsolutePath());
            }
        } else if (e.getSource() == chooseDestinationButton) {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
            int option = fileChooser.showOpenDialog(this);
            if (option == JFileChooser.APPROVE_OPTION) {
                File file = fileChooser.getSelectedFile();
                destinationTextField.setText(file.getAbsolutePath() + "\\");
            }
        } else if (e.getSource() == generateButton) {
            String srcPath = sourceTextField.getText();
            String destinationPath = destinationTextField.getText();
            if ("".equals(srcPath) || srcPath == null || destinationPath == null || destinationPath.equals("")) {
                System.out.println("bad paths");
            } else {
                ExcelReader excelReader = new ExcelReader(srcPath);
                List<List<DocumentModel>> allTables = excelReader.getAllTables();
                drawErrorMessages();

                new ExcelWriter().writeDocument(allTables, destinationPath);
                drawErrorMessages();
            }
        }
    }

    public static void main(String[] args) {
        new MainWindow();

//        ge.luka.ExcelReader excelReader = new ge.luka.ExcelReader(
//                "C:\\Users\\lukaa\\Downloads\\გაყვანები სამშაბათი ხუთშაბათი შაბათი კლინიკების მიხედვით.xlsx");
//        List<List<DocumentModel>> allTables = excelReader.getAllTables();
////        for (int i = 0; i < allTables.size(); i++) {
////            System.out.println(allTables.get(i));
////            System.out.println("end of " + i + "th driver");
////        }
//
//        ge.luka.ExcelWriter excelWriter = new ge.luka.ExcelWriter();
//        excelWriter.writeDocument(allTables, "C:\\Users\\lukaa\\Downloads\\test.xlsx");

    }
}
