# Print
打印报表到excel中


Private Sub DoEveryStationMemberCardYear()
        Try
            Dim year As String = ddlYear.SelectedItem.Value
            
            Dim thisBegin As DateTime = MCO001.YMDChar2Date(year + "0101")
            Dim thisEnd As DateTime = thisBegin.AddYears(1)
            thisBegin = thisBegin.AddDays(-1)
            
            Dim thisBeginDate As String = thisBegin.ToString("yyyy-MM-dd")
            Dim thisEndDate As String = thisEnd.ToString("yyyy-MM-dd")

            
            Dim lastBegin As DateTime = thisBegin.AddYears(-1)
            Dim lastEnd As DateTime = thisEnd.AddYears(-1)

            
            Dim lastBeginDate As String = lastBegin.ToString("yyyy-MM-dd")
            Dim lastEndDate As String = lastEnd.ToString("yyyy-MM-dd")

            Dim stationID As String = Trim(ddlStation.SelectedItem.Value)
          
            Dim everyStationMemberCardYear As DataTable = APICurrent.API.Dynamic.HuiLeCards.getEveryStationMemberCardYear(stationID, thisBeginDate, thisEndDate)
            
            Dim everyStationMemberCardYearToMonth As DataTable = APICurrent.API.Dynamic.HuiLeCards.getEveryStationMemberCardYearToMonth(stationID, thisBeginDate, thisEndDate)
         
            Dim lastEveryStationMemberCardYear As DataTable = APICurrent.API.Dynamic.HuiLeCards.getEveryStationMemberCardYear(stationID, lastBeginDate, lastEndDate)
            
            Dim lastEveryStationMemberCardYearToMonth As DataTable = APICurrent.API.Dynamic.HuiLeCards.getEveryStationMemberCardYearToMonth(stationID, lastBeginDate, lastEndDate)
            
            Dim countStation As DataTable = APICurrent.API.Dynamic.HuiLeCards.getStationCount(stationID, thisBeginDate, thisEndDate, lastBeginDate, lastEndDate)
            
            Dim countStationMX As DataTable = APICurrent.API.Dynamic.HuiLeCards.getStationCountMX(stationID, thisBeginDate, thisEndDate, lastBeginDate, lastEndDate)
            Dim path As String = ExportEveryStationMemberCardYear(countStation, countStationMX, everyStationMemberCardYear, everyStationMemberCardYearToMonth, lastEveryStationMemberCardYear, lastEveryStationMemberCardYearToMonth)

            If Not String.IsNullOrEmpty(path) Then
                DownLoadFile(path)
            Else
                Me.Controls.Add(New LiteralControl("<SCRIPT LANGUAGE=JavaScript>alert('" & MCO003.getNameByResID(MCO005.IMSG011) & "')</script>"))
            End If
        Catch ex As Exception
            Me.Controls.Add(New LiteralControl("<SCRIPT LANGUAGE=JavaScript>alert('" & MCO003.getNameByResID(MCO005.IMSG011) & "')</script>"))
        Finally
            GC.Collect()
        End Try
    End Sub
	
	
	
	
	
	
	
	Private Function ExportEveryStationMemberCardYear(ByVal stationData As DataTable, _
                                                      ByVal stationDataMX As DataTable, _
                                                      ByVal everyStationMemberCareYear As DataTable, _
                                                      ByVal everyStationMemberCardYearToMonth As DataTable, _
                                                      ByVal lastEveryStationMemberCardYear As DataTable, _
                                                      ByVal lastEveryStationMemberCardYearToMonth As DataTable) As String
        Dim objExcelType As Type
        Dim excel As Object
        Dim workbook As Object
        Dim worksheet As Object
        Dim path As String
        Try
            objExcelType = Type.GetTypeFromProgID("Excel.Application", True)
            excel = System.Activator.CreateInstance(objExcelType)
            excel.Visible = False
            Dim rootPath As String = Server.MapPath("~")
            rootPath = rootPath + "\\Report\\EveryStationMemberCardYear.xlsx"
            workbook = excel.Workbooks.Open(rootPath)
            worksheet = workbook.Worksheets(1)
            Dim row As Integer = 6
            Dim column As Integer = 0
            Dim thisCount As Integer = everyStationMemberCareYear.Rows.Count        '数据的总数
            Dim thisDataCount As Integer = everyStationMemberCardYearToMonth.Rows.Count    '数据明细的总数

            Dim lastCount As Integer = lastEveryStationMemberCardYear.Rows.Count    '的数据的总数
            Dim lastDataCount As Integer = lastEveryStationMemberCardYearToMonth.Rows.Count    
			'数据明细的总数

            Dim countStationMX As Integer = stationDataMX.Rows.Count     '服务站的数量-明细
            Dim totalCountMX As Integer = countStationMX * 12            '明细里面的总行数


            Dim countStation As Integer = stationData.Rows.Count        '


            
            Dim thisCareMoney As Integer    
            Dim thisPutInMoney As Integer   
            Dim thisTimeBank As Integer     

            Dim lastCareMoney As Integer    
            Dim lastPutInMoney As Integer   
            Dim lastTimeBank As Integer     

            Dim diffCareMoney As Integer    
            Dim diffPutInMoney As Integer   
            Dim diffTimeBank As Integer     

            Dim thisCareMoneyMX As Integer    
            Dim thisPutInMoneyMX As Integer   
            Dim thisTimeBankMX As Integer     

            Dim lastCareMoneyMX As Integer    
            Dim lastPutInMoneyMX As Integer   
            Dim lastTimeBankMX As Integer     

            Dim diffCareMoneyMX As Integer    
            Dim diffPutInMoneyMX As Integer   
            Dim diffTimeBankXM As Integer     

            Dim totalThis As Integer
            Dim totalLast As Integer
            Dim totalDiff As Integer


            Dim totalThisMX As Integer
            Dim totalLastMX As Integer
            Dim totalDiffMX As Integer


            Dim year As String = ddlYear.SelectedItem.Text
            'newData里面的2个数字，分别表示报表里面的总行数和总列数                        
            Dim thisData(thisCount, 14) As Object    '今年的数据             
            Dim thisDataMX(totalCountMX + thisCount + 9, 14) As Object    '今年的数据明细

            Dim lastData(lastCount, 13) As Object    '去年的数据
            Dim lastDateMX(lastDataCount, 13) As Object  '去年的数据明细

            Dim dataTitel(2, 1) As Object
            '如果行或者列合并了,则赋值的时候，将模板上面的合并的地方先取消合并，看看值是具体写在哪里，然后在该位置上面赋值
            dataTitel(0, 0) = "*****"
            dataTitel(1, 0) = "年份:" + year + "年"

            '==============
            Dim i As Integer
            Dim l As Integer
            Dim m As Integer
            Dim j As Integer
            Dim n As Integer
            Dim sName As String
            Dim nDate As String
            Dim accountTypeID As String

            For m = 0 To countStation - 1
                Dim stationName As String = MCO001.DBNULL2Str(stationData.Rows(m)(HuiLeRechargeField.STATIONNAME_FIELD))

                thisData(m, column) = stationName
                For i = 0 To thisCount - 1

                    worksheet.Range(worksheet.Cells(i + row, 1), worksheet.Cells(i + row, 2)).Merge()
                    worksheet.Range(worksheet.Cells(i + row, 1), worksheet.Cells(i + row, 2)).HorizontalAlignment = 3

                    sName = MCO001.DBNULL2Str(everyStationMemberCareYear.Rows(i)(HuiLeRechargeField.STATIONNAME_FIELD))

                    If (sName.Equals(stationName)) Then
                        thisData(i, column) = MCO001.DBNULL2Str(everyStationMemberCareYear.Rows(i)(HuiLeRechargeField.STATIONNAME_FIELD))
                        thisData(i, column + 2) = MCO001.DBNULL2Str(everyStationMemberCareYear.Rows(i)(HuiLeRechargeField.CAREMONEY_FIELD))
                        thisCareMoney += MCO001.DBNULL2Int(thisData(i, column + 2))

                        thisData(i, column + 5) = MCO001.DBNULL2Str(everyStationMemberCareYear.Rows(i)(HuiLeRechargeField.PUTINMONEY_FIELD))
                        thisPutInMoney += MCO001.DBNULL2Int(thisData(i, column + 5))

                        thisData(i, column + 8) = MCO001.DBNULL2Str(everyStationMemberCareYear.Rows(i)(HuiLeRechargeField.TIMEBANK_FIELD))
                        thisTimeBank += MCO001.DBNULL2Int(thisData(i, column + 8))

                        thisData(i, column + 11) = MCO001.DBNULL2Str(MCO001.DBNULL2Int(thisData(i, column + 2)) + MCO001.DBNULL2Int(thisData(i, column + 5)))
                        totalThis += MCO001.DBNULL2Int(thisData(i, column + 11))
                    End If


                Next

                For n = 0 To lastCount - 1
                    sName = MCO001.DBNULL2Str(lastEveryStationMemberCardYear.Rows(n).Item(HuiLeRechargeField.STATIONNAME_FIELD))

                    If (sName.Equals(stationName)) Then

                        thisData(i, column + 3) = MCO001.DBNULL2Str(lastEveryStationMemberCardYear.Rows(n)(HuiLeRechargeField.CAREMONEY_FIELD))
                        lastCareMoney += MCO001.DBNULL2Int(thisData(i, column + 3))

                        thisData(i, column + 6) = MCO001.DBNULL2Str(lastEveryStationMemberCardYear.Rows(n)(HuiLeRechargeField.PUTINMONEY_FIELD))
                        lastPutInMoney += MCO001.DBNULL2Int(thisData(i, column + 6))

                        thisData(i, column + 9) = MCO001.DBNULL2Str(lastEveryStationMemberCardYear.Rows(n)(HuiLeRechargeField.TIMEBANK_FIELD))
                        lastTimeBank += MCO001.DBNULL2Int(thisData(i, column + 9))

                        thisData(i, column + 12) = MCO001.DBNULL2Str(MCO001.DBNULL2Int(thisData(i, column + 3)) + MCO001.DBNULL2Int(thisData(i, column + 6)))
                        totalLast += MCO001.DBNULL2Int(thisData(i, column + 12))
                    End If

                Next
                thisData(i, column + 4) = MCO001.DBNULL2Int(thisData(i, column + 2)) - MCO001.DBNULL2Int(thisData(i, column + 4))
                thisData(i, column + 7) = MCO001.DBNULL2Int(thisData(i, column + 5)) - MCO001.DBNULL2Int(thisData(i, column + 6))
                thisData(i, column + 10) = MCO001.DBNULL2Int(thisData(i, column + 8)) - MCO001.DBNULL2Int(thisData(i, column + 9))
            Next


            worksheet.Range(worksheet.Cells(thisCount + row, 1), worksheet.Cells(thisCount + row, 2)).Merge()
            worksheet.Range(worksheet.Cells(thisCount + row, 1), worksheet.Cells(thisCount + row, 2)).HorizontalAlignment = 3
            thisData(thisCount, column) = "*****"
            thisData(thisCount, column + 2) = thisCareMoney
            thisData(thisCount, column + 3) = lastCareMoney
            thisData(thisCount, column + 4) = MCO001.DBNULL2Int(thisCareMoney) - MCO001.DBNULL2Int(lastCareMoney)
            thisData(thisCount, column + 5) = thisPutInMoney
            thisData(thisCount, column + 6) = lastPutInMoney
            thisData(thisCount, column + 7) = MCO001.DBNULL2Int(thisPutInMoney) - MCO001.DBNULL2Int(lastPutInMoney)
            thisData(thisCount, column + 8) = thisTimeBank
            thisData(thisCount, column + 9) = lastTimeBank
            thisData(thisCount, column + 10) = MCO001.DBNULL2Int(thisTimeBank) - MCO001.DBNULL2Int(lastTimeBank)
            thisData(thisCount, column + 11) = totalThis
            thisData(thisCount, column + 12) = totalLast
            thisData(thisCount, column + 13) = MCO001.DBNULL2Int(totalThis) - MCO001.DBNULL2Int(totalLast)

            '==============

            '==============


            thisDataMX(0, 0) = "*****"
            worksheet.Range(worksheet.Cells(thisCount + row + 2, 1), worksheet.Cells(thisCount + row + 4, 1)).Merge()
            worksheet.Range(worksheet.Cells(thisCount + row + 2, 1), worksheet.Cells(thisCount + row + 4, 1)).HorizontalAlignment = 3
            thisDataMX(0, 1) = "*****"
            worksheet.Range(worksheet.Cells(thisCount + row + 2, 2), worksheet.Cells(thisCount + row + 4, 2)).Merge()
            worksheet.Range(worksheet.Cells(thisCount + row + 2, 2), worksheet.Cells(thisCount + row + 4, 2)).HorizontalAlignment = 3
            thisDataMX(0, 2) = "*****"
            worksheet.Range(worksheet.Cells(thisCount + row + 2, 3), worksheet.Cells(thisCount + row + 2, 11)).Merge()
            worksheet.Range(worksheet.Cells(thisCount + row + 2, 3), worksheet.Cells(thisCount + row + 2, 11)).HorizontalAlignment = 3
            thisDataMX(0, 11) = "*****"
            worksheet.Range(worksheet.Cells(thisCount + row + 2, 12), worksheet.Cells(thisCount + row + 3, 14)).Merge()
            worksheet.Range(worksheet.Cells(thisCount + row + 2, 12), worksheet.Cells(thisCount + row + 3, 14)).HorizontalAlignment = 3

            thisDataMX(1, 2) = "*****"
            worksheet.Range(worksheet.Cells(thisCount + row + 3, 3), worksheet.Cells(thisCount + row + 3, 5)).Merge()
            worksheet.Range(worksheet.Cells(thisCount + row + 3, 3), worksheet.Cells(thisCount + row + 3, 5)).HorizontalAlignment = 3
            thisDataMX(1, 5) = "*****"
            worksheet.Range(worksheet.Cells(thisCount + row + 3, 6), worksheet.Cells(thisCount + row + 3, 8)).Merge()
            worksheet.Range(worksheet.Cells(thisCount + row + 3, 6), worksheet.Cells(thisCount + row + 3, 8)).HorizontalAlignment = 3
            thisDataMX(1, 8) = "*****"
            worksheet.Range(worksheet.Cells(thisCount + row + 3, 9), worksheet.Cells(thisCount + row + 3, 11)).Merge()
            worksheet.Range(worksheet.Cells(thisCount + row + 3, 9), worksheet.Cells(thisCount + row + 3, 11)).HorizontalAlignment = 3

            thisDataMX(2, 2) = "*****"
            thisDataMX(2, 3) = "*****"
            thisDataMX(2, 4) = "*****"

            thisDataMX(2, 5) = "*****"
            thisDataMX(2, 6) = "*****"
            thisDataMX(2, 7) = "*****"

            thisDataMX(2, 8) = "*****"
            thisDataMX(2, 9) = "*****"
            thisDataMX(2, 10) = "*****"

            thisDataMX(2, 11) = "*****"
            thisDataMX(2, 12) = "*****"
            thisDataMX(2, 13) = "*****"

            For m = 0 To countStationMX - 1
                Dim stationName As String = MCO001.DBNULL2Str(stationDataMX.Rows(m)(HuiLeRechargeField.STATIONNAME_FIELD))
                For l = 1 To 12
                    Dim dateFormat As String = IIf(l < 10, "0" + l.ToString(), l.ToString())
                    thisDataMX((m) * 12 + 2 + l, column) = stationName
                    thisDataMX((m) * 12 + 2 + l, column + 1) = dateFormat

                    '数据明细
                    For j = 0 To thisDataCount - 1
                        sName = MCO001.DBNULL2Str(everyStationMemberCardYearToMonth.Rows(j)(HuiLeRechargeField.STATIONNAME_FIELD))
                        nDate = MCO001.DBNULL2Str(everyStationMemberCardYearToMonth.Rows(j)(HuiLeRechargeField.HUILEUSEDATE_FIELD))
                        accountTypeID = MCO001.DBNULL2Str(everyStationMemberCardYearToMonth.Rows(j)(HuiLeRechargeField.ACCOUNTTYPEID_FIELD))
                        If (sName.Equals(stationName) And nDate.Equals(dateFormat) And MCO002.AccountTypeID_01.Equals(accountTypeID)) Then
                            thisDataMX((m) * 12 + 2 + l, column + 2) = MCO001.DBNULL2Str(everyStationMemberCardYearToMonth.Rows(j)(HuiLeRechargeField.CAREMONEY_FIELD))
                            thisCareMoneyMX += MCO001.DBNULL2Int(thisDataMX((m) * 12 + 2 + l, column + 2))
                        End If
                        If (sName.Equals(stationName) And nDate.Equals(dateFormat) And MCO002.AccountTypeID_02.Equals(accountTypeID)) Then
                            thisDataMX((m) * 12 + 2 + l, column + 5) = MCO001.DBNULL2Str(everyStationMemberCardYearToMonth.Rows(j)(HuiLeRechargeField.PUTINMONEY_FIELD))
                            thisPutInMoneyMX += MCO001.DBNULL2Int(thisDataMX((m) * 12 + 2 + l, column + 5))
                        End If
                        If (sName.Equals(stationName) And nDate.Equals(dateFormat) And MCO002.AccountTypeID_03.Equals(accountTypeID)) Then
                            thisDataMX((m) * 12 + 2 + l, column + 8) = MCO001.DBNULL2Str(everyStationMemberCardYearToMonth.Rows(j)(HuiLeRechargeField.TIMEBANK_FIELD))
                            thisTimeBankMX += MCO001.DBNULL2Int(thisDataMX((m) * 12 + 2 + l, column + 8))
                        End If
                        
                        thisDataMX((m) * 12 + 2 + l, column + 11) = MCO001.DBNULL2Str(MCO001.DBNULL2Int(thisDataMX((m) * 12 + 2 + l, column + 2)) + MCO001.DBNULL2Int(thisDataMX((m) * 12 + 2 + l, column + 5)))
                    Next

                    For n = 0 To lastDataCount - 1
                        sName = MCO001.DBNULL2Str(lastEveryStationMemberCardYearToMonth.Rows(n)(HuiLeRechargeField.STATIONNAME_FIELD))
                        nDate = MCO001.DBNULL2Str(lastEveryStationMemberCardYearToMonth.Rows(n)(HuiLeRechargeField.HUILEUSEDATE_FIELD))                        
                        If (sName.Equals(stationName) And nDate.Equals(dateFormat)) Then
                            thisDataMX((m) * 12 + 2 + l, column + 3) = MCO001.DBNULL2Str(lastEveryStationMemberCardYearToMonth.Rows(n)(HuiLeRechargeField.CAREMONEY_FIELD))
                            lastCareMoney += MCO001.DBNULL2Int(thisDataMX((m) * 12 + 2 + l, column + 3))

                            thisDataMX((m) * 12 + 2 + l, column + 6) = MCO001.DBNULL2Str(lastEveryStationMemberCardYearToMonth.Rows(n)(HuiLeRechargeField.PUTINMONEY_FIELD))
                            lastPutInMoneyMX += MCO001.DBNULL2Int(thisDataMX((m) * 12 + 2 + l, column + 6))

                            thisDataMX((m) * 12 + 2 + l, column + 9) = MCO001.DBNULL2Str(lastEveryStationMemberCardYearToMonth.Rows(n)(HuiLeRechargeField.TIMEBANK_FIELD))
                            lastTimeBankMX += MCO001.DBNULL2Int(thisDataMX((m) * 12 + 2 + l, column + 9))

                            thisDataMX((m) * 12 + 2 + l, column + 12) = MCO001.DBNULL2Str(MCO001.DBNULL2Int(thisDataMX((m) * 12 + 2 + l, column + 3)) + MCO001.DBNULL2Int(thisDataMX((m) * 12 + 2 + l, column + 6)))
                            totalLastMX += MCO001.DBNULL2Int(thisDataMX((m) * 12 + 2 + l, column + 12))
                        End If
                    Next
                    thisDataMX((m) * 12 + 2 + l, column + 4) = MCO001.DBNULL2Int(thisDataMX((m) * 12 + 2 + l, column + 2)) - MCO001.DBNULL2Int(thisDataMX((m) * 12 + 2 + l, column + 3))
                    thisDataMX((m) * 12 + 2 + l, column + 7) = MCO001.DBNULL2Int(thisDataMX((m) * 12 + 2 + l, column + 5)) - MCO001.DBNULL2Int(thisDataMX((m) * 12 + 2 + l, column + 6))
                    thisDataMX((m) * 12 + 2 + l, column + 10) = MCO001.DBNULL2Int(thisDataMX((m) * 12 + 2 + l, column + 8)) - MCO001.DBNULL2Int(thisDataMX((m) * 12 + 2 + l, column + 9))
                Next
            Next


            thisDataMX(totalCountMX + 3, column) = "合计"
            worksheet.Range(worksheet.Cells(thisCount + row + totalCountMX + 5, 1), worksheet.Cells(thisCount + row + totalCountMX + 5, 2)).Merge()
            worksheet.Range(worksheet.Cells(thisCount + row + totalCountMX + 5, 1), worksheet.Cells(thisCount + row + totalCountMX + 5, 2)).HorizontalAlignment = 3
            thisDataMX(totalCountMX + 3, column + 2) = thisCareMoneyMX
            thisDataMX(totalCountMX + 3, column + 3) = lastCareMoneyMX
            thisDataMX(totalCountMX + 3, column + 4) = MCO001.DBNULL2Int(thisCareMoneyMX) - MCO001.DBNULL2Int(lastCareMoneyMX)
            thisDataMX(totalCountMX + 3, column + 5) = thisPutInMoneyMX
            thisDataMX(totalCountMX + 3, column + 6) = lastPutInMoneyMX
            thisDataMX(totalCountMX + 3, column + 7) = MCO001.DBNULL2Int(thisPutInMoneyMX) - MCO001.DBNULL2Int(lastPutInMoneyMX)
            thisDataMX(totalCountMX + 3, column + 8) = thisTimeBankMX
            thisDataMX(totalCountMX + 3, column + 9) = lastTimeBankMX
            thisDataMX(totalCountMX + 3, column + 10) = MCO001.DBNULL2Int(thisTimeBankMX) - MCO001.DBNULL2Int(lastTimeBankMX)
            thisDataMX(totalCountMX + 3, column + 11) = MCO001.DBNULL2Int(thisCareMoneyMX) + MCO001.DBNULL2Int(thisPutInMoneyMX)
            thisDataMX(totalCountMX + 3, column + 12) = MCO001.DBNULL2Int(lastCareMoneyMX) + MCO001.DBNULL2Int(lastPutInMoneyMX)
            thisDataMX(totalCountMX + 3, column + 13) = MCO001.DBNULL2Int(thisDataMX(totalCountMX + 3, column + 11)) - MCO001.DBNULL2Int(thisDataMX(totalCountMX + 3, column + 12))

            '==============数据明细结束

            Dim T As Object = worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(2, 1))
            T.Value = dataTitel

            '数据的表格样式
            Dim thisBorderRange As Object = worksheet.Range(worksheet.Cells(row, 1), worksheet.Cells(row + thisCount, column + 14))
            thisBorderRange.Borders.LineStyle = 1

            '数据明细的表格样式            
            Dim thisBorderRangeMX As Object = worksheet.Range(worksheet.Cells(thisCount + row + 2, 1), worksheet.Cells(thisCount + row + totalCountMX + 5, column + 14))
            thisBorderRangeMX.Borders.LineStyle = 1

            '第一个worksheet.Cells表示数据的起始位置，今年的数据的起始位置
            Dim tDate As Object = worksheet.Range(worksheet.Cells(row, 1), worksheet.Cells(row + thisCount, column + 14))
            tDate.Value = thisData

            '数据明细的起始位置
            Dim lData As Object = worksheet.Range(worksheet.Cells(thisCount + row + 2, 1), worksheet.Cells(thisCount + row + totalCountMX + 5, column + 14))
            lData.Value = thisDataMX

            excel.DisplayAlerts = False
            path = CreatePath("各服务站会员卡年报汇总.xlsx")
            workbook.SaveAs(path)
        Catch ex As Exception
            path = Nothing
            LogManager.Error(ex.ToString())
            Throw
        Finally
            workbook.Close()
            excel.Quit()
            worksheet = Nothing
            workbook = Nothing
            excel = Nothing
        End Try
        Return path
    End Function
