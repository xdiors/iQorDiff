Set objExcel = CreateObject("Excel.Application")
objExcel.Workbooks.Open ("F:\zyz\smart\QorDiff.xlsm")
objExcel.Run("ģ��1.sub1")
objExcel.Workbooks.Close()