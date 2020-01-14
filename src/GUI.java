import javax.swing.*;
import java.awt.*;

public class GUI extends JFrame
{
    JPanel panel = new JPanel();
    JPanel showPanel = new JPanel();
    public GUI()
    {
        setTitle("근태 포맷 변경 프로그램");
        setSize(900,600);
        add(panel);
        init();
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setVisible(true);
    }

    public void init()
    {
        GridLayout gl = new GridLayout(2,1);
        JButton excelLoading = new JButton("엑셀 불러오기");
        JButton excelSave = new JButton("엑셀 저장하기");
        panel.add(excelLoading);
        panel.add(excelSave);
        panel.add(showPanel);
        ExcelButton listener = new ExcelButton();
        excelLoading.addActionListener(listener);
        excelSave.addActionListener(listener);
    }


}
