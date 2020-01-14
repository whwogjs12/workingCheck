import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

public class ExcelButton implements ActionListener {

    @Override
    public void actionPerformed(ActionEvent e) {
        JButton selectedButton = (JButton)e.getSource();
        if(selectedButton.getText().equals("엑셀 불러오기"))
        {
            JFileChooser fcl = new JFileChooser(new File("c:\\"));
            fcl.setDialogTitle("파일 불러오기");
            fcl.setFileFilter(new FilexlsxFilter("xlsx", "Xlsx File"));
            int result = fcl.showOpenDialog(null);
            if(result == JFileChooser.APPROVE_OPTION)
            {

            }
        }
        else
        {
            JFileChooser fcs = new JFileChooser(new File("c:\\"));
            fcs.setDialogTitle("파일 저장하기");
            fcs.setFileFilter(new FilexlsxFilter("xlsx", "Xlsx File"));
            int result = fcs.showSaveDialog(null);
            if(result == JFileChooser.APPROVE_OPTION)
            {

            }
        }
    }
}
