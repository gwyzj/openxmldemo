   IWorkbook workbook = new XSSFWorkbook();
   ISheet sheet = workbook.CreateSheet("扫描数据");            //在工作簿中：建立空白工作表
   IRow row = sheet.CreateRow(0);                              //在工作表中：建立行，参数为行号，从0计
   ICell cell = row.CreateCell(0);                             //在行中：建立单元格，参数为列号，从0计
   cell.SetCellValue("扫描数据");              //设置单元格内容
     //新建一个字体样式对象

   ICellStyle style = workbook.CreateCellStyle();
   style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center; //设置单元格的样式：水平对齐居中
   style.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center; //设置单元格的样式：水平对齐居中
   IFont font = workbook.CreateFont();                         //新建一个字体样式对象
                                                               //font.Boldweight = short.MaxValue;                           //设置字体加粗样式
   font.FontName = "宋体";
   font.FontHeight = 16;
   style.SetFont(font);                                        //使用SetFont方法将字体样式添加到单元格样式中 
   cell.CellStyle = style;                                     //将新的样式赋给单元格
   row.Height = 30 * 20;                                       //设置单元格的高度
   NPOI.SS.Util.CellRangeAddress region = new NPOI.SS.Util.CellRangeAddress(0, 0, 0, 15);//设置一个合并单元格区域，使用上下左右定义CellRangeAddress区域
   sheet.AddMergedRegion(region);

   IRow row1 = sheet.CreateRow(1);
   row1.CreateCell(0).SetCellValue("实际值");
   row1.CreateCell(1).SetCellValue("最大值");
   row1.CreateCell(2).SetCellValue("最小值");
   row1.CreateCell(3).SetCellValue("平均值");

   row1.CreateCell(4).SetCellValue(" ");
   row1.CreateCell(5).SetCellValue("最大偏差");
   row1.CreateCell(6).SetCellValue("最小偏差");
   row1.CreateCell(7).SetCellValue("平均值偏差");
   row1.CreateCell(8).SetCellValue("离散度");
   row1.CreateCell(9).SetCellValue("平均差");
   row1.CreateCell(10).SetCellValue("标准差");

   row1.CreateCell(11).SetCellValue(" ");
   row1.CreateCell(12).SetCellValue("绝对最大偏差");
   row1.CreateCell(13).SetCellValue("绝对最小偏差");
   row1.CreateCell(14).SetCellValue("绝对平均偏差");
   row1.CreateCell(15).SetCellValue("偏移量");
   row1.CreateCell(16).SetCellValue("绝对精度");
   //设置单元格的宽度
   for (int i = 0; i < 15; i++)
   {
       sheet.SetColumnWidth(i, 13 * 256);//设置单元格的宽度
   }

   Recordlist = Recordlist.OrderByDescending(p => p.m_int32readlisc).ToList();
   int len = Recordlist.Count;
   int temp = 0;

   int splitrowindex = 0;

   if (len > 0)
   {
       for (int i = 0; i < len; i++)
       {
           IRow rowx = sheet.CreateRow(i + 2);
           temp = i;
           if (((RecordStruct)Recordlist[temp]).m_int32ave < splitvalue)
           {
               if (splitrowindex == 0)
               {
                   splitrowindex = i + 2;
               }

           }
           rowx.CreateCell(0).SetCellValue(((RecordStruct)Recordlist[temp]).m_int32readlisc);
           rowx.CreateCell(1).SetCellValue(((RecordStruct)Recordlist[temp]).m_int32Max);
           rowx.CreateCell(2).SetCellValue(((RecordStruct)Recordlist[temp]).m_int32Min);
           rowx.CreateCell(3).SetCellValue(((RecordStruct)Recordlist[temp]).m_int32ave);
           rowx.CreateCell(4);
           rowx.CreateCell(5);
           rowx.CreateCell(6);
           rowx.CreateCell(7);
           rowx.CreateCell(8);
           rowx.CreateCell(9);
           rowx.CreateCell(10);
           //rowx.CreateCell(9).SetCellValue(((RecordStruct)Recordlist[temp]).m_MAD);
           //rowx.CreateCell(10).SetCellValue(((RecordStruct)Recordlist[temp]).m_stdDeviation);
           rowx.CreateCell(11);
           rowx.CreateCell(12);
           rowx.CreateCell(13);
           rowx.CreateCell(14);
           rowx.CreateCell(15);
           rowx.CreateCell(16);
           rowx.CreateCell(17);
           rowx.CreateCell(18);
       }
       //设置保留小数点个数
       ICellStyle cellStyle = workbook.CreateCellStyle();
       cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");

       //先远距离偏差
       IRow row2 = sheet.GetRow(2);
       ICell cellx13 = row2.GetCell(15);
       // cellx13.SetCellFormula("AVERAGE(H3:H" + (len + 2).ToString() + ")"); //偏移量：平均偏差的平均值
       cellx13.SetCellFormula("AVERAGE(H3:H" + splitrowindex + ")"); //偏移量：平均偏差的平均值
       cellx13.CellStyle = cellStyle;
       cellx13 = row2.GetCell(18);
       cellx13.SetCellFormula($"P3+{main.remote_offset}");
       cellx13.CellStyle = cellStyle;
       //后近距离偏差
       IRow row21 = sheet.GetRow(3);
       ICell cellxnear = row21.GetCell(15);
       // cellx13.SetCellFormula("AVERAGE(H3:H" + (len + 2).ToString() + ")"); //偏移量：平均偏差的平均值
       cellxnear.SetCellFormula($"AVERAGE(H{splitrowindex + 1}:H{len + 2})"); //偏移量：平均偏差的平均值
       cellxnear.CellStyle = cellStyle;
       cellxnear = row21.GetCell(18);
       cellxnear.SetCellFormula($"P4+{main.near_offset}");
       cellxnear.CellStyle = cellStyle;


       if (len > 0)
       {
           for (int i = 0; i < len; i++)
           {
               IRow rowx = sheet.GetRow(i + 2);
               string rowno = (i + 3).ToString();
               ICell cellx5 = rowx.GetCell(5);
               cellx5.SetCellFormula("SUM(B" + rowno + ",A" + rowno + "*(-1))");  //最大偏差： 最大值减实际值

               ICell cellx6 = rowx.GetCell(6);
               cellx6.SetCellFormula("SUM(C" + rowno + ",A" + rowno + "*(-1))"); //最小偏差：最小值减实际值

               ICell cellx7 = rowx.GetCell(7);
               cellx7.SetCellFormula("SUM(D" + rowno + ",A" + rowno + "*(-1))"); //平均偏差： 平均值减实际值

               ICell cellx8 = rowx.GetCell(8);
               cellx8.SetCellFormula("SUM(B" + rowno + ",C" + rowno + "*(-1))"); //离散度：最大值减最小值

               if (i + 2 < splitrowindex)
               {
                   ICell cellx10 = rowx.GetCell(12);
                   cellx10.SetCellFormula("SUM(F" + rowno + ",P3*(-1))");  //绝对最大偏差：最大偏差减去偏移量
                   cellx10.CellStyle = cellStyle;

                   ICell cellx11 = rowx.GetCell(13);
                   cellx11.SetCellFormula("SUM(G" + rowno + ",P3*(-1))"); //绝对最小偏差：最小偏差减去偏移量
                   cellx11.CellStyle = cellStyle;

                   ICell cellx12 = rowx.GetCell(14);
                   cellx12.SetCellFormula("SUM(H" + rowno + ",P3*(-1))"); //绝对平均偏差：平均偏差减去偏移量
                   cellx12.CellStyle = cellStyle;
               }
               else
               {
                   ICell cellx10 = rowx.GetCell(12);
                   cellx10.SetCellFormula("SUM(F" + rowno + ",P4*(-1))");  //绝对最大偏差：最大偏差减去偏移量
                   cellx10.CellStyle = cellStyle;

                   ICell cellx11 = rowx.GetCell(13);
                   cellx11.SetCellFormula("SUM(G" + rowno + ",P4*(-1))"); //绝对最小偏差：最小偏差减去偏移量
                   cellx11.CellStyle = cellStyle;

                   ICell cellx12 = rowx.GetCell(14);
                   cellx12.SetCellFormula("SUM(H" + rowno + ",P4*(-1))"); //绝对平均偏差：平均偏差减去偏移量
                   cellx12.CellStyle = cellStyle;
               }
           }
       }

       //绝对精度
       IRow row5 = sheet.GetRow(2);
       ICell cell514 = row5.GetCell(16);
       cell514.SetCellFormula($"SUM(MAX(F3:F{splitrowindex}),MIN(G3:G{splitrowindex})*(-1))");
       cell514 = row5.GetCell(17);
       cell514.SetCellValue("远距离");



       IRow row6 = sheet.GetRow(3);
       ICell cell614 = row6.GetCell(16);
       cell614.SetCellFormula($"SUM(MAX(F{splitrowindex + 1}:F{len + 2}),MIN(G{splitrowindex + 1}:G{len + 2})*(-1))");
       cell614 = row6.GetCell(17);
       cell614.SetCellValue("近距离");

       IRow row7 = sheet.GetRow(4);
       ICell cell714 = row7.GetCell(16);
       cell714.SetCellFormula($"SUM(MAX(M3:M{len + 2}),MIN(N3:N{len + 2})*(-1))");
       cell714 = row7.GetCell(17);
       cell714.SetCellValue("整体");

       path = Environment.CurrentDirectory.ToString() + @"\Excel\"  + ToolFunc.g_smac+ DateTime.Now.ToString("-yyyy-MM-dd-HH-mm-ss") + ".xlsx";
       Directory.CreateDirectory(Environment.CurrentDirectory + @"\Excel\");
       FileStream fs = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
       if (fs == null)
       {
           return;
       }
       workbook.Write(fs);
       workbook.Close();
       fs.Close();
       //MessageBox.Show("数据保存成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
       //加载excel
       System.Diagnostics.Process.Start(path);
       
   }
