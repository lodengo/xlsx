xlsx
====
目的：读写excel/xlsx 文件
  
环境：linux+nodejs
  
考虑：
  1. 不想依赖MS Office
  2. csv格式对多sheet、格式化、数据图表支持有限
  3. 不想调用强大的PHPExcel
  4. 使用zip压缩/解压+xml操作来读写xlsx
  5. 复杂excel报表导出先做好模板上传服务器，服务器只填充数据导出
  
参考：  
  https://github.com/SheetJS/js-xlsx  
  https://github.com/node-xmpp/ltx


