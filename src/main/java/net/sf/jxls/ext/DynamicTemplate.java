package net.sf.jxls.ext;

import net.sf.jxls.transformer.XLSTransformer;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author: wang.q
 * <code>
 *     List list = new ArrayList<DynamicColumnModel>();
       list.add(new DynamicColumnModel("网盟", "yougou_union"));
       list.add(new DynamicColumnModel("王钦", "wq"));
       list.add(new DynamicColumnModel("日期", "create_date", FormatConstant.DATE_LONG_CN);
       java.io.InputStream inputStream = DynamicTemplate.generate(list);
 * </code>
 */
public class DynamicTemplate {
    private static Log log = LogFactory.getLog(DynamicTemplate.class);

    public static InputStream generate(List<DynamicColumnModel> dynamicColumns) {
        log.info("Begin Generate:Jxls-Excel_Template");
        Map beans = new HashMap();
        beans.put("data", dynamicColumns);

        XLSTransformer transformer = new XLSTransformer(new DynamicConfiguration());
        String dynamicTemplate = "dynamicColumnTemplate.xls";
        Resource resource = new ClassPathResource(dynamicTemplate);
        try {
            Workbook workbook = transformer.transformXLS(resource.getInputStream(), beans);

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            workbook.write(out);
            log.info("End :Jxls-Excel_Template...");
            //导出本地
            //workbook = transformer.transformXLS(resource.getInputStream(), beans);
            //workbook.write(new FileOutputStream("d:\\asdf.xls"));

            return new ByteArrayInputStream(out.toByteArray());
        } catch (InvalidFormatException e) {
            log.error("Generate: Jxls-Excel_Template failed，invalid format.");
            e.printStackTrace();
        } catch (IOException e) {
            log.error("Generate: Jxls-Excel_Template，IO exception.");
            e.printStackTrace();
        }

        return new ByteArrayInputStream(new byte[]{});
    }
}
