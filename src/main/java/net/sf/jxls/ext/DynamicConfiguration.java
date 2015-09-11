package net.sf.jxls.ext;

/**
 * @author: wang.q
 */
import net.sf.jxls.tag.JxTaglib;
import net.sf.jxls.tag.TagLib;
import net.sf.jxls.transformer.Configuration;
import org.apache.commons.digester.Digester;

import java.util.*;

public class DynamicConfiguration extends Configuration {
    private String startExpressionToken = "#(";
    private String endExpressionToken = ")";

    private String tagPrefix = "dynjx:";
    private String tagPrefixWithBrace = "<dynjx:";
    private String forTagName = "forEach";
    private String forTagItems = "items";
    private String forTagVar = "var";

    private HashMap tagLibs = new HashMap();
    private Digester digester;
    private String jxlsRoot;
    private boolean encodeXMLAttributes = true;
    String sheetKeyName = "sheet";
    String workbookKeyName = "workbook";
    String rowKeyName = "hssfRow";
    String cellKeyName = "hssfCell";

    private String excludeSheetProcessingMark = "#Exclude";
    boolean removeExcludeSheetProcessingMark = false;
    Set excludeSheets = new HashSet();

    public DynamicConfiguration() {
        registerTagLib(new JxTaglib(), "dynjx");
    }

    public String getTagPrefix() {
        return tagPrefix;
    }

    public void setTagPrefix(String tagPrefix) {
        this.tagPrefix = tagPrefix;
    }

    public String getTagPrefixWithBrace() {
        return tagPrefixWithBrace;
    }

    public void setTagPrefixWithBrace(String tagPrefixWithBrace) {
        this.tagPrefixWithBrace = tagPrefixWithBrace;
    }

    public String getForTagName() {
        return forTagName;
    }

    public void setForTagName(String forTagName) {
        this.forTagName = forTagName;
    }

    public String getForTagItems() {
        return forTagItems;
    }

    public void setForTagItems(String forTagItems) {
        this.forTagItems = forTagItems;
    }

    public String getForTagVar() {
        return forTagVar;
    }

    public void setForTagVar(String forTagVar) {
        this.forTagVar = forTagVar;
    }

    public String getStartExpressionToken() {
        return startExpressionToken;
    }

    public void setStartExpressionToken(String startExpressionToken) {
        this.startExpressionToken = startExpressionToken;
    }

    public String getEndExpressionToken() {
        return endExpressionToken;
    }

    public void setEndExpressionToken(String endExpressionToken) {
        this.endExpressionToken = endExpressionToken;
    }

    public static final String NAMESPACE_URI = "http://jxls.sourceforge.net/jxls";
    public static final String JXLS_ROOT_TAG = "dynjxls";
    public static final String JXLS_ROOT_START = "<dynjx:dynjxls xmlns:jx=\"" + NAMESPACE_URI + "\">";
    public static final String JXLS_ROOT_END = "</dynjx:dynjxls>";

    public void registerTagLib(TagLib tagLib, String namespace) {

        if (null == this.tagLibs) {
            this.tagLibs = new HashMap();
        }
        if (this.tagLibs.containsKey(namespace)) {
            throw new RuntimeException("Duplicate tag-lib namespace: " + namespace);
        }

        this.tagLibs.put(namespace, tagLib);
    }

    public Digester getDigester() {

        synchronized (this) {
            if (digester == null) {
                initDigester();
            }
        }

        return digester;
    }

    private void initDigester() {
        digester = new Digester();
        digester.setNamespaceAware(true);
        digester.setValidating(false);

        StringBuffer sb = new StringBuffer();
        sb.append("<dynjxls ");
        boolean firstTime = true;

        Map.Entry entry = null;

        for (Iterator itr = tagLibs.entrySet().iterator(); itr.hasNext();) {

            entry = (Map.Entry) itr.next();

            String namespace = (String) entry.getKey();
            String namespaceURI = DynamicConfiguration.NAMESPACE_URI + "/" + namespace;
            digester.setRuleNamespaceURI(namespaceURI);

            if (firstTime) {
                firstTime = false;
            } else {
                sb.append(" ");
            }
            sb.append("xmlns:");
            sb.append(namespace);
            sb.append("=\"");
            sb.append(namespaceURI);
            sb.append("\"");

            TagLib tagLib = (TagLib) entry.getValue();
            Map.Entry tagEntry = null;
            for (Iterator itr2 = tagLib.getTags().entrySet().iterator(); itr2.hasNext();) {
                tagEntry = (Map.Entry) itr2.next();
                digester.addObjectCreate(DynamicConfiguration.JXLS_ROOT_TAG + "/" + tagEntry.getKey(), (String) tagEntry.getValue());
                digester.addSetProperties(DynamicConfiguration.JXLS_ROOT_TAG + "/" + tagEntry.getKey());
            }
        }

        sb.append(">");
        this.jxlsRoot = sb.toString();
    }

    public String getJXLSRoot() {

        synchronized(this) {
            if (jxlsRoot == null) {
                initDigester();
            }
        }

        return jxlsRoot;
    }

    public Set getExcludeSheets() {
        return this.excludeSheets;
    }

    public void addExcludeSheet(String name) {
        this.excludeSheets.add(name);
    }

    public String getJXLSRootEnd() {
        return "</dynjxls>";
    }

    public boolean getEncodeXMLAttributes() {
        return encodeXMLAttributes;
    }

    public void setEncodeXMLAttributes(boolean encodeXMLAttributes) {
        this.encodeXMLAttributes = encodeXMLAttributes;
    }
}
