using System.Data;

namespace AchieveItSystemBLL
{
    public interface IExcelService
    {
        DataTable GetExcelData(string excelPath);
        void SaveDataToExcel(string excelPath, DataTable data, int selectedRowIndex, int completedColumnIndex, int completionDateColumnIndex, int overdueColumnIndex);
        void AddNewTaskToExcel(string excelPath, string date, string theme, string task, bool isCompleted);
        void CreateExcelTemplate(string filePath);
    }
}