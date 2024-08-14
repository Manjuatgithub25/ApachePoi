def readExcelTable(filePath, sheetName) {
    FileInputStream file = new FileInputStream(new File(filePath))
    Workbook workbook = new XSSFWorkbook(file)
    Sheet sheet = workbook.getSheet(sheetName)

    def tableHtml = "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse: collapse;'>"
    for (Row row : sheet) {
        tableHtml += "<tr>"
        for (Cell cell : row) {
            String cellValue = ""
            switch (cell.cellType) {
                case CellType.STRING:
                    cellValue = cell.stringCellValue
                    break
                case CellType.NUMERIC:
                    cellValue = cell.numericCellValue.toString()
                    break
                case CellType.BOOLEAN:
                    cellValue = cell.booleanCellValue.toString()
                    break
                default:
                    cellValue = ""
            }
            tableHtml += "<td>${cellValue}</td>"
        }
        tableHtml += "</tr>"
    }
    tableHtml += "</table>"

    file.close()
    return tableHtml
}
