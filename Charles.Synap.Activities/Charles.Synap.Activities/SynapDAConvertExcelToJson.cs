using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using Newtonsoft.Json;

namespace Charles.Synap.Activities
{
    public class SynapDAConvertExcelToJson : CodeActivity
    {
        [Category("Input")]
        public InArgument<DataTable> InputDataTable { get; set; }
        
        [Category("Input")]
        public InArgument<string> OutputJsonFilePath { get; set; }
        
        [Category("Output")]
        public OutArgument<int> RecordCount { get; set; }
        
        [Category("Output")]
        public OutArgument<string> ErrorMessage { get; set; }

        private int recordCount;

        public SynapDAConvertExcelToJson()
        {
            recordCount = 0;
        }

        private bool ConvertDataTableToJson(DataTable dataTable, string jsonFilePath)
        {
#if DEBUG
            //Debugger.Launch();
#endif
            
            try
            {
                // 출력 폴더가 없으면 생성
                string outputDirectory = Path.GetDirectoryName(jsonFilePath);
                if (!string.IsNullOrEmpty(outputDirectory) && !Directory.Exists(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                }

                var jsonData = ConvertDataTableToJsonData(dataTable);
                
                // JSON 파일로 저장
                string jsonString = JsonConvert.SerializeObject(jsonData, Formatting.Indented);
                File.WriteAllText(jsonFilePath, jsonString, Encoding.UTF8);
                
                recordCount = jsonData.Count;
                
#if DEBUG
                Console.WriteLine($"DataTable을 JSON으로 변환 완료: {recordCount}개 레코드");
#endif
            }
            catch (Exception ex)
            {
                Console.WriteLine($"오류 발생: {ex.Message}");
                recordCount = -1;
                return false;
            }
            
            return true;
        }

        private List<Dictionary<string, object>> ConvertDataTableToJsonData(DataTable dataTable)
        {
            var result = new List<Dictionary<string, object>>();
            
            if (dataTable == null || dataTable.Rows.Count == 0)
                return result;

            foreach (DataRow row in dataTable.Rows)
            {
                var record = new Dictionary<string, object>();
                bool hasData = false;

                foreach (DataColumn column in dataTable.Columns)
                {
                    object cellValue = row[column];
                    
                    // DBNull을 null로 변환
                    if (cellValue == DBNull.Value)
                        cellValue = null;
                    
                    if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                    {
                        hasData = true;
                    }

                    record[column.ColumnName] = cellValue;
                }

                if (hasData)
                {
                    result.Add(record);
                }
            }

            return result;
        }

        protected override void Execute(CodeActivityContext context)
        {
            var dataTable = InputDataTable.Get(context);
            var jsonFilePath = OutputJsonFilePath.Get(context);

            if (dataTable == null)
            {
                ErrorMessage.Set(context, "Input DataTable is null");
                RecordCount.Set(context, 0);
                return;
            }

            if (string.IsNullOrEmpty(jsonFilePath))
            {
                ErrorMessage.Set(context, "Output JSON file path is empty");
                RecordCount.Set(context, 0);
                return;
            }

            if (ConvertDataTableToJson(dataTable, jsonFilePath))
            {
                RecordCount.Set(context, recordCount);
                ErrorMessage.Set(context, string.Empty);
            }
            else
            {
                ErrorMessage.Set(context, "Error in SynapDAConvertExcelToJson");
                RecordCount.Set(context, 0);
            }
        }
    }
}