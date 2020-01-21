import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableCellRenderer;
import javax.swing.table.TableColumnModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import org.apache.poi.xssf.usermodel.*;


public class GUI extends JFrame implements ActionListener {
    JPanel panel = new JPanel(new BorderLayout());
    JPanel excelPanel = new JPanel();
    JPanel showPanel = new JPanel();
    WorkingData wd = new WorkingData();

    public GUI() {
        setTitle("근태 포맷 변경 프로그램");
        setSize(1200, 600);
        add(panel);
        init();
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setVisible(true);
    }

    public void init() {
        JButton excelLoading = new JButton("엑셀 불러오기");
        JButton excelSave = new JButton("엑셀 저장하기");
        excelPanel.add(excelLoading);
        excelPanel.add(excelSave);
        panel.add(excelPanel, BorderLayout.NORTH);
        panel.add(showPanel, BorderLayout.CENTER);
        excelLoading.addActionListener(this);
        excelSave.addActionListener(this);
    }


    @Override
    public void actionPerformed(ActionEvent e) {
        JButton selectedButton = (JButton) e.getSource();
        JFileChooser fc = new JFileChooser(new File("c:\\"));
        fc.setMultiSelectionEnabled(false);
        fc.setFileFilter(new FilexlsxFilter("xlsx", "excel File"));
        int result;
        String filePath = null;
        if (selectedButton.getText().equals("엑셀 불러오기")) {
            fc.setDialogTitle("파일 불러오기");
            result = fc.showOpenDialog(null);
            if (result == JFileChooser.APPROVE_OPTION) {
                filePath = fc.getSelectedFile().toString();
                try {
                    dataTransfer(filePath);
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
            }
        } else {
            fc.setDialogTitle("파일 저장하기");
            result = fc.showSaveDialog(null);
            if (result == JFileChooser.APPROVE_OPTION) {
                filePath = fc.getSelectedFile().toString();
            }
        }
        System.out.println(filePath);

    }

    public void dataTransfer(String filePath) throws IOException {
        FileInputStream fis = new FileInputStream(filePath);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        ArrayList<ArrayList<String>> list = new ArrayList<>();
        XSSFSheet sheet = workbook.getSheetAt(0);
        int rowIndex = sheet.getPhysicalNumberOfRows();
        wd = new WorkingData();
        wd.createAttribute(sheet, 0);//첫번째 속성
        wd.createAttribute(sheet, 1);//두번째 속성
        wd.createPersonData(sheet, rowIndex);

    }

    public void dataSave(String filePath) throws IOException
    {

    }
}
