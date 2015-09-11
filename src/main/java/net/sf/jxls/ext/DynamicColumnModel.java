package net.sf.jxls.ext;

/**
 * @author: wang.q
 */
import java.io.Serializable;

public final class DynamicColumnModel implements Serializable {
    private static final long serialVersionUID = 3076049038299609232L;

    private String headerName;
    private String colExp;
    // Excel 单元格格式 自定义列值
    private String formatCode = null;

    public DynamicColumnModel(String headerName, String colExp) {
        this.headerName = headerName;
        this.colExp = colExp;
    }

    public DynamicColumnModel(String headerName, String colExp, String formatCode) {
        this.headerName = headerName;
        this.colExp = colExp;
        this.formatCode = formatCode;
    }

    public String getHeaderName() {
        return headerName;
    }

    public void setHeaderName(String headerName) {
        this.headerName = headerName;
    }

    public String getColExp() {
        return colExp;
    }

    public void setColExp(String colExp) {
        this.colExp = colExp;
    }

    public String getFormatCode() {
        return formatCode;
    }

    public void setFormatCode(String formatCode) {
        this.formatCode = formatCode;
    }

    @Override
    public int hashCode() {
        int prime = 997;
        int headerNameHashCode = prime + (null == headerName ? 0 : headerName.hashCode());
        int colExpHashCode = prime + (null == colExp ? 0 : colExp.hashCode());
        return prime + headerNameHashCode + colExpHashCode;
    }

    @Override
    public boolean equals(Object obj) {
        if (null == obj) {
            return false;
        }

        if (this == obj) {
            return true;
        }

        if (!obj.getClass().equals(this.getClass())) {
            return false;
        }

        DynamicColumnModel other = (DynamicColumnModel)obj;
        return other.hashCode() == this.hashCode();
    }
}
