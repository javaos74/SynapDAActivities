using System;
using System.Data;
using System.IO;

namespace Charles.Synap.Activities
{
    // 테스트용 간단한 클래스 (실제 사용시에는 제거)
    public class TestDataTableToJson
    {
        public static DataTable CreateSampleDataTable()
        {
            var dataTable = new DataTable("학생정보");
            
            // 컬럼 추가
            dataTable.Columns.Add("이름", typeof(string));
            dataTable.Columns.Add("나이", typeof(int));
            dataTable.Columns.Add("수학점수", typeof(int));
            dataTable.Columns.Add("영어점수", typeof(int));
            
            // 데이터 추가
            dataTable.Rows.Add("김철수", 25, 85, 90);
            dataTable.Rows.Add("이영희", 23, 92, 88);
            dataTable.Rows.Add("박민수", 24, 78, 85);
            dataTable.Rows.Add("최지영", 22, 95, 92);
            
            return dataTable;
        }
        
        public static void TestConversion()
        {
            var activity = new SynapDAConvertExcelToJson();
            var dataTable = CreateSampleDataTable();
            string outputPath = Path.Combine(Path.GetTempPath(), "test_output.json");
            
            // 실제 테스트는 UiPath 환경에서 수행해야 함
            Console.WriteLine($"테스트 DataTable 생성 완료: {dataTable.Rows.Count}개 행");
            Console.WriteLine($"출력 경로: {outputPath}");
        }
    }
}