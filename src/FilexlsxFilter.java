import javax.swing.filechooser.FileFilter;
import java.io.File;
import java.io.FilenameFilter;
import java.nio.file.spi.FileTypeDetector;

public class FilexlsxFilter extends FileFilter {

    private String extention;
    private String description;

    public FilexlsxFilter(String extention, String description)
    {
        this.extention = extention;
        this.description = description;
    }

    @Override
    public boolean accept(File f) {
        if(f.isDirectory() || f.getName().endsWith(extention)) return true;
        return false;
    }

    @Override
    public String getDescription() {
        return null;
    }
}
