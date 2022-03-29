package cn.gjing.excel.style.lock;

import cn.gjing.excel.base.exception.ExcelException;
import cn.gjing.excel.base.listener.write.ExcelSheetWriteListener;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Excel Sheet lock listener, used to lock all cells in the Sheet to prevent user changes
 *
 * @author Gjing
 **/
public class ExcelSheetLockListener implements ExcelSheetWriteListener {
    /**
     * Lock password
     */
    private final String password;

    public ExcelSheetLockListener() {
        this.password = "abc123";
    }

    public ExcelSheetLockListener(String password) {
        if (password == null || password.isEmpty()) {
            throw new ExcelException("lock password cannot be empty");
        }
        this.password = password;
    }

    @Override
    public void completeSheet(Sheet sheet) {
        sheet.protectSheet(this.password);
    }
}
