package net.sf.jxls.ext;

import net.sf.jxls.ext.util.StringUtils;
import net.sf.jxls.parser.Cell;
import net.sf.jxls.parser.Expression;
import net.sf.jxls.processor.CellProcessor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author: wang.q
 */
public class DataFormatCellProcessor implements CellProcessor {
    private List<DynamicColumnModel> cols = new ArrayList<DynamicColumnModel>();
    Map<String, CellStyle> cellStyles = new HashMap<String, CellStyle>();

    public DataFormatCellProcessor(List<DynamicColumnModel> cols) {
        this.cols = cols;
    }

    @Override
    public void processCell(Cell cell, Map namedCells) {
        try {
            if (cell.getExpressions().size() > 0) {
                Expression expression = (Expression) cell.getExpressions().get(0);
                if (StringUtils.isNotBlank(expression.toString())) {
                    String currDataIndex = expression.toString().substring(2);
                    for (int i=0; i< cols.size(); i++) {
                        DynamicColumnModel dyn = cols.get(i);
                        if (dyn.getColExp().equalsIgnoreCase(currDataIndex)
                                && StringUtils.isNotBlank(dyn.getFormatCode())) {

                            String currKey = String.format("dyn_%s", currDataIndex);
                            CellStyle cellStyle = duplicateStyle(currKey, dyn.getFormatCode(), cell);
                            cell.getPoiCell().setCellStyle(cellStyle);

                            break;
                        }
                    }
                }
               /*
                Property property = (Property) expression.getProperties().get(0);
                if (property != null && property.getBeanName() != null && property.getBeanName().indexOf(beanName) >= 0 && property.getBean() instanceof DynamicColumnModel) {
                    DynamicColumnModel col = (DynamicColumnModel) property.getBean();
                    if (StringUtils.isNotBlank(col.getFormatCode())) {
                        Workbook workbook = cell.getPoiCell().getSheet().getWorkbook();
                        CellStyle cellStyle = workbook.createCellStyle();
                        DataFormat format = workbook.createDataFormat();
                        cellStyle.setDataFormat(format.getFormat(col.getFormatCode()));
                        cell.getPoiCell().setCellStyle(cellStyle);
                    }
                }
                */
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private CellStyle duplicateStyle(String key, String dynFormatCode, Cell cell){
        if(cellStyles.containsKey(key) && null != cellStyles.get(key)){
            return cellStyles.get(key);
        }

        CellStyle style;
        try {
            Workbook workbook = cell.getPoiCell().getSheet().getWorkbook();

            style = workbook.createCellStyle();
            DataFormat format = workbook.createDataFormat();
            style.setDataFormat(format.getFormat(dynFormatCode));
            cellStyles.put(key, style);
        } catch (Exception e) {
            e.printStackTrace();
            style = cell.getPoiCell().getCellStyle();
        }

        return style;
    }
}
