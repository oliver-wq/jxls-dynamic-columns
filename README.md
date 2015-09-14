# jxls-dynamic-columns
dynamic export

// ctor

List<DynamicColumnModel> dynCols = extractDynamicCols(jsonGridCols);
InputStream jxlsTemplate = DynamicTemplate.generate(dynCols);

map.addAttribute("DynamicCols", dynCols);
map.addAttribute("DynamicTemplate", jxlsTemplate);

// export

XLSTransformer dynaTrans = new XLSTransformer();
dynaTrans.registerCellProcessor(new DataFormatCellProcessor((List<DynamicColumnModel>)map.get("DynamicCols")));
Workbook book = dynaTrans.transformXLS((java.io.InputStream)map.get("DynamicTemplate"), map);

// Servlet Output

response.setContentType(...);
response.setHeader(...);
ServletOutputStream out = response.getOutputStream();
book.write(out);
out.flush();
out.close();
