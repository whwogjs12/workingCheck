import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableCellRenderer;
import javax.swing.table.TableColumnModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;


public class GUI extends JFrame implements ActionListener {
    JPanel panel = new JPanel(new BorderLayout());
    JPanel excelPanel = new JPanel();
    JPanel showPanel = new JPanel();
    WorkingData wd = new WorkingData();

    public GUI() {
        setTitle("근태 포맷 변경 프로그램");
        setSize(400, 80);
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
                    JOptionPane.showMessageDialog(null,"파일을 불러오는데 성공했습니다", "파일 데이터 전달 완료",JOptionPane.INFORMATION_MESSAGE);
                } catch (IOException e1) {
                    e1.printStackTrace();
                    JOptionPane.showMessageDialog(null,"파일을 불러오는데 실패했습니다", "파일 에러",JOptionPane.ERROR_MESSAGE);
                }
            }
        } else {
            fc.setDialogTitle("파일 저장하기");
            result = fc.showSaveDialog(null);
            if (result == JFileChooser.APPROVE_OPTION) {
                filePath = fc.getSelectedFile().toString();
                try {
                    dataSave(filePath);
                    JOptionPane.showMessageDialog(null,"파일을 저장하는데 성공했습니다", "파일 저장 완료",JOptionPane.INFORMATION_MESSAGE);
                } catch (IOException e1) {
                    e1.printStackTrace();
                    JOptionPane.showMessageDialog(null,"문제가 발생하였습니다.", "파일 저장 실패",JOptionPane.INFORMATION_MESSAGE);
                }
            }
        }
        System.out.println(filePath);

    }

    public void dataTransfer(String filePath) throws IOException { //엑셀 불러오기
        FileInputStream fis = new FileInputStream(filePath);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheetAt(0);
        int rowIndex = sheet.getPhysicalNumberOfRows();
        wd = new WorkingData();
        wd.createAttribute(sheet, 0);//첫번째 속성
        wd.createAttribute(sheet, 1);//두번째 속성
        wd.createPersonData(sheet, rowIndex);

    }

    public void dataSave(String filePath) throws IOException //엑셀 저장하기
    {
        String fileName = "근태현황";
        XSSFWorkbook saveBook = new XSSFWorkbook();
        XSSFSheet sheet = saveBook.createSheet(fileName);
        attributeEntry(sheet);
        for(int x=0;x<wd.getList().size();x++)
        {
            dataEntry(sheet,wd.getList().get(x),x*6+2);
        }
        File file = new File(filePath + ".xlsx");
        FileOutputStream fos = new FileOutputStream(file);
        saveBook.write(fos);
        fos.close();
    }
    public void attributeEntry(XSSFSheet sheet)
    {
        XSSFRow row = sheet.createRow(0);
        XSSFCell cell = row.createCell(1);
        ArrayList<String> dateList = wd.getDateList();
        ArrayList<String> weekList = wd.getWeekList();
        cell.setCellValue("성명");
        cell = row.createCell(2);
        cell.setCellValue("구분");
        cell = row.createCell(3);
        cell.setCellValue("계");
        for(int x = 2;x<dateList.size()-1;x++)
        {
            cell = row.createCell(x+2);
            cell.setCellValue(dateList.get(x));
        }
        row = sheet.createRow(1);
        for(int y = 2;y<weekList.size();y++)
        {
            cell = row.createCell(y+2);
            cell.setCellValue(weekList.get(y));
        }
        sheet.addMergedRegion(new CellRangeAddress(0,1,1,1));
        sheet.addMergedRegion(new CellRangeAddress(0,1,2,2));
        sheet.addMergedRegion(new CellRangeAddress(0,1,3,3));
    }

    public void dataEntry(XSSFSheet sheet, ArrayList<String> data, int cellIndex) // 데이터 출력
    {
        XSSFRow row = null;
        XSSFCell cell = null;
        for(int x= 0;x<6;x++)
        {
            row = sheet.createRow(cellIndex+x);
            cell = row.createCell(2);
            String divisionMessage = divisionMessageString(x);
            cell.setCellValue(divisionMessage);
        }
        sheet.addMergedRegion(new CellRangeAddress(cellIndex,cellIndex+5,0,0));
        sheet.addMergedRegion(new CellRangeAddress(cellIndex,cellIndex+5,1,1));
        row = sheet.getRow(cellIndex);
        cell = row.createCell(0); // 성명 파트
        cell.setCellValue(data.get(0));
        cell = row.createCell(1); //수 파트
        cell.setCellValue(data.get(1));

        for(int i=2;i<data.size()-1;i++)
        {
            ArrayList<String> weekList = wd.getWeekList();
            try
            {
                int dataParts = Integer.parseInt(data.get(i));
                if(weekList.get(i).equals("토")) // 주말의 경우는 전부 특근
                {
                    row = sheet.getRow(cellIndex);
                    cell = row.createCell(i+2);
                    cell.setCellValue(0.0);
                    row = sheet.getRow(cellIndex+4); // cellIndex + 4 는 특근
                    cell = row.createCell(i+2);
                    cell.setCellValue(dataParts);
                }
                else if( weekList.get(i).equals("일"))
                {
                    row = sheet.getRow(cellIndex+4); // cellIndex + 4 는 특근
                    cell = row.createCell(i+2);
                    cell.setCellValue(dataParts);
                    row = sheet.getRow(cellIndex);
                    if(weeklyMeasurement(row, i+2))
                    {
                        cell = row.createCell(i+2);
                        cell.setCellValue(1.0);
                    }

                }
                else
                {
                    row = sheet.getRow(cellIndex); // cellIndex는 출근
                    cell = row.createCell(i+2);
                    cell.setCellValue(1.0);
                    if(dataParts > 8)
                    {
                        if(dataParts >= 12)
                        {
                            row = sheet.getRow(cellIndex+3);
                            cell = row.createCell(i+2);
                            cell.setCellValue(4.0);
                            row = sheet.getRow(cellIndex+5);
                            cell = row.createCell(i+2);
                            cell.setCellValue(8.0);
                        }
                        else
                        {
                            row = sheet.getRow(cellIndex+3);
                            cell = row.createCell(i+2);
                            cell.setCellValue(dataParts - 8.0);
                        }
                    }
                }
            }
            catch (NumberFormatException e)
            {
                row = sheet.getRow(cellIndex);
                String holiday = data.get(i);
                switch (holiday)
                {
                    case "연":
                        cell = row.createCell(i+2);
                        cell.setCellValue(1.0);
                        row = sheet.getRow(cellIndex+1);
                        holiday = "연차";
                        break;
                    case "무":
                        holiday = "무휴";
                        break;
                    case "결":
                        holiday = "결근";
                        break;
                }
                if(holiday.equals("휴") && weekList.get(i).equals("일"))
                {
                    row = sheet.getRow(cellIndex);
                    if(weeklyMeasurement(row, i+2))
                    {
                        cell = row.createCell(i+2);
                        cell.setCellValue(1.0);
                    }
                    else
                    {
                        cell = row.createCell(i+2);
                        cell.setCellValue(0.0);
                    }
                }
                else if (holiday.equals("휴") && weekList.get(i).equals("토"))
                {
                    row = sheet.getRow(cellIndex);
                    cell = row.createCell(i+2);
                    cell.setCellValue(0.0);
                }
                else
                {
                    cell = row.createCell(i+2);
                    cell.setCellValue(holiday);
                }
            }
        }
    }

    public String divisionMessageString (int rowIndex) //구분 칸 채우기
    {
        String divisionMessage =null;
        switch (rowIndex) {
            case 0:
                divisionMessage = "출근";
                break;
            case 1:
                divisionMessage = "지각";
                break;
            case 2:
                divisionMessage = "조퇴";
                break;
            case 3:
                divisionMessage = "잔업";
                break;
            case 4:
                divisionMessage = "특근";
                break;
            case 5:
                divisionMessage = "야간";
                break;
        }
        return divisionMessage;
    }

    /*public String changedForm (String dataToChange, int dataIndex)
    {
        try {
            int data = Integer.parseInt(dataToChange);
            ArrayList<String> weekList = wd.getWeekList();
            if(weekList.get(dataIndex).equals("토") || weekList.get(dataIndex).equals("일"))
            {

            }
        }
        catch (NumberFormatException e)
        {
            return dataToChange;
        }
    }*/

    public boolean weeklyMeasurement(XSSFRow row, int column)
    {
        XSSFCell cell= null;
        if(column <= 10)
        {
            return false;
        }
        else
        {
            for(int a= column-2;a>=column-6;a--)
            {
                cell = row.getCell(a);
                String check = null;
                if(cell !=null)
                {
                    switch (cell.getCellType())
                    {
                        case NUMERIC:
                            check = cell.getNumericCellValue() +"";
                            break;
                        case STRING:
                            check = cell.getStringCellValue();
                    }

                    if(check.equals("1") || check.equals("1.0")){continue;}
                    else {return false;}
                }
                else
                {
                    return false;
                }
            }
        }
        return true;
    }
}
