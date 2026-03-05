using Microsoft.Data.SqlClient;
using Microsoft.Win32;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Windows;

namespace WpfLab3;

public partial class MainWindow : Window
{
    // ----------------------------------------------------------------
    // LocalDB — база создаётся автоматически рядом с .exe
    // Не нужен SSMS или SQL Server
    // ----------------------------------------------------------------
    private static readonly string DbPath =
        Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Lab3.mdf");

    private static readonly string ConnectionString =
        $@"Data Source=(LocalDB)\MSSQLLocalDB;
           AttachDbFilename={DbPath};
           Integrated Security=True;
           Connect Timeout=30;";

    public MainWindow()
    {
        InitializeComponent();
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        InitDatabase();   // создаём БД и таблицу при первом запуске
        LoadGrid();
    }

    // ================================================================
    //  ИНИЦИАЛИЗАЦИЯ БАЗЫ ДАННЫХ (создаётся автоматически)
    // ================================================================
    private void InitDatabase()
    {
        try
        {
            // Если .mdf уже существует — просто подключаемся
            if (!File.Exists(DbPath))
            {
                // Создаём новый файл БД через master
                string masterConn =
                    @"Data Source=(LocalDB)\MSSQLLocalDB;
                      Initial Catalog=master;
                      Integrated Security=True;";

                using var conn = new SqlConnection(masterConn);
                conn.Open();

                string createDb = $@"
                    CREATE DATABASE Lab3
                    ON PRIMARY (NAME='Lab3', FILENAME='{DbPath}')
                    LOG ON  (NAME='Lab3_log',
                             FILENAME='{DbPath.Replace(".mdf", "_log.ldf")}')";

                using var cmd = new SqlCommand(createDb, conn);
                cmd.ExecuteNonQuery();
            }

            // Создаём таблицу если её нет
            using var c = new SqlConnection(ConnectionString);
            c.Open();

            const string createTable = @"
                IF NOT EXISTS (
                    SELECT 1 FROM sys.tables WHERE name = 'Employees'
                )
                CREATE TABLE Employees (
                    Id       INT IDENTITY(1,1) PRIMARY KEY,
                    Login    NVARCHAR(100) NOT NULL,
                    Password NVARCHAR(256) NOT NULL,
                    Role     NVARCHAR(100) NOT NULL,
                    FullName NVARCHAR(200) NULL,
                    Email    NVARCHAR(150) NULL
                )";

            using var ct = new SqlCommand(createTable, c);
            ct.ExecuteNonQuery();

            SetStatus("✅ База данных готова");
        }
        catch (Exception ex)
        {
            SetStatus($"❌ Ошибка БД: {ex.Message}");
            MessageBox.Show($"Ошибка при инициализации БД:\n\n{ex.Message}",
                            "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // ================================================================
    //  ИМПОРТ из 5.xlsx
    // ================================================================
    private void BtnImport_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Title  = "Выберите файл 5.xlsx",
            Filter = "Excel файлы (*.xlsx)|*.xlsx"
        };
        if (dlg.ShowDialog() != true) return;

        try
        {
            var employees = ReadExcel(dlg.FileName);
            SaveToDb(employees);
            LoadGrid();
            SetStatus($"✅ Импортировано {employees.Count} записей из «{Path.GetFileName(dlg.FileName)}»");
            MessageBox.Show($"Импорт успешно завершён.\nЗаписей добавлено: {employees.Count}",
                            "Импорт", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            SetStatus("❌ Ошибка импорта");
            MessageBox.Show($"Ошибка при импорте:\n\n{ex.Message}",
                            "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private List<Employee> ReadExcel(string path)
    {
        var list = new List<Employee>();

        using var pkg = new ExcelPackage(new FileInfo(path));
        var ws = pkg.Workbook.Worksheets[0];
        int rows = ws.Dimension?.Rows ?? 0;

        // Строка 1 — заголовки; данные со строки 2
        for (int r = 2; r <= rows; r++)
        {
            string login = ws.Cells[r, 1].Text.Trim();
            if (string.IsNullOrWhiteSpace(login)) continue;

            list.Add(new Employee
            {
                Login    = login,
                Password = Sha256(ws.Cells[r, 2].Text.Trim()),
                Role     = ws.Cells[r, 3].Text.Trim(),
                FullName = ws.Cells[r, 4].Text.Trim(),
                Email    = ws.Cells[r, 5].Text.Trim()
            });
        }
        return list;
    }

    private void SaveToDb(List<Employee> list)
    {
        using var conn = new SqlConnection(ConnectionString);
        conn.Open();

        // Очищаем таблицу перед импортом
        new SqlCommand("DELETE FROM Employees", conn).ExecuteNonQuery();

        const string sql = @"
            INSERT INTO Employees (Login, Password, Role, FullName, Email)
            VALUES (@Login, @Password, @Role, @FullName, @Email)";

        foreach (var emp in list)
        {
            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@Login",    emp.Login);
            cmd.Parameters.AddWithValue("@Password", emp.Password);
            cmd.Parameters.AddWithValue("@Role",     emp.Role);
            cmd.Parameters.AddWithValue("@FullName", string.IsNullOrEmpty(emp.FullName) ? DBNull.Value : emp.FullName);
            cmd.Parameters.AddWithValue("@Email",    string.IsNullOrEmpty(emp.Email)    ? DBNull.Value : emp.Email);
            cmd.ExecuteNonQuery();
        }
    }

    // ================================================================
    //  ЭКСПОРТ в Excel — группировка по роли
    // ================================================================
    private void BtnExport_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new SaveFileDialog
        {
            Title    = "Сохранить экспорт",
            Filter   = "Excel файлы (*.xlsx)|*.xlsx",
            FileName = $"Export_ByRole_{DateTime.Now:yyyyMMdd_HHmm}.xlsx"
        };
        if (dlg.ShowDialog() != true) return;

        try
        {
            var all = GetAllFromDb();
            if (all.Count == 0)
            {
                MessageBox.Show("Нет данных для экспорта. Сначала выполните импорт.",
                                "Нет данных", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            ExportByRole(all, dlg.FileName);
            SetStatus($"✅ Экспортировано в «{Path.GetFileName(dlg.FileName)}»");
            MessageBox.Show("Экспорт завершён успешно!", "Экспорт",
                            MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            SetStatus("❌ Ошибка экспорта");
            MessageBox.Show($"Ошибка при экспорте:\n\n{ex.Message}",
                            "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ExportByRole(List<Employee> employees, string filePath)
    {
        using var pkg = new ExcelPackage();

        var groups = employees
            .GroupBy(e => string.IsNullOrWhiteSpace(e.Role) ? "Без роли" : e.Role)
            .OrderBy(g => g.Key);

        foreach (var grp in groups)
        {
            string name = grp.Key.Length > 31 ? grp.Key[..31] : grp.Key;
            var ws = pkg.Workbook.Worksheets.Add(name);

            // --- Заголовок группы ---
            ws.Cells[1, 1].Value = $"Роль: {grp.Key}";
            ws.Cells[1, 1, 1, 2].Merge = true;
            ws.Cells[1, 1, 1, 2].Style.Font.Bold = true;
            ws.Cells[1, 1, 1, 2].Style.Font.Size = 13;
            ws.Cells[1, 1, 1, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            SetCellColor(ws.Cells[1, 1, 1, 2], 44, 62, 80, 255, 255, 255);

            // --- Заголовки столбцов ---
            ws.Cells[2, 1].Value = "Логин";
            ws.Cells[2, 2].Value = "Пароль (SHA-256)";
            ws.Cells[2, 1, 2, 2].Style.Font.Bold = true;
            SetCellColor(ws.Cells[2, 1, 2, 2], 41, 128, 185, 255, 255, 255);

            // --- Данные ---
            int row = 3;
            foreach (var emp in grp)
            {
                ws.Cells[row, 1].Value = emp.Login;
                ws.Cells[row, 2].Value = emp.Password;

                if (row % 2 == 0)
                    SetCellColor(ws.Cells[row, 1, row, 2], 236, 240, 241, 0, 0, 0);

                row++;
            }

            ws.Cells[ws.Dimension.Address].AutoFitColumns();
        }

        pkg.SaveAs(new FileInfo(filePath));
    }

    // ================================================================
    //  ВСПОМОГАТЕЛЬНЫЕ
    // ================================================================
    private void BtnRefresh_Click(object sender, RoutedEventArgs e)
    {
        LoadGrid();
        SetStatus("🔄 Данные обновлены");
    }

    private void LoadGrid()
    {
        try
        {
            var data = GetAllFromDb();
            dgEmployees.ItemsSource = data;
            SetStatus($"📊 Записей в БД: {data.Count}");
        }
        catch (Exception ex)
        {
            SetStatus($"⚠️ Ошибка загрузки: {ex.Message}");
        }
    }

    private List<Employee> GetAllFromDb()
    {
        var list = new List<Employee>();
        using var conn = new SqlConnection(ConnectionString);
        conn.Open();

        using var cmd = new SqlCommand(
            "SELECT Id, Login, Password, Role, FullName, Email FROM Employees ORDER BY Role, Login",
            conn);
        using var r = cmd.ExecuteReader();
        while (r.Read())
            list.Add(new Employee
            {
                Id       = r.GetInt32(0),
                Login    = r.GetString(1),
                Password = r.GetString(2),
                Role     = r.IsDBNull(3) ? "" : r.GetString(3),
                FullName = r.IsDBNull(4) ? "" : r.GetString(4),
                Email    = r.IsDBNull(5) ? "" : r.GetString(5)
            });

        return list;
    }

    private static string Sha256(string text)
    {
        if (string.IsNullOrEmpty(text)) return "";
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(text));
        return Convert.ToHexString(bytes).ToLowerInvariant();
    }

    private static void SetCellColor(ExcelRange range,
        int r, int g, int b, int fr, int fg, int fb)
    {
        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
        range.Style.Fill.BackgroundColor.SetColor(
            System.Drawing.Color.FromArgb(r, g, b));
        range.Style.Font.Color.SetColor(
            System.Drawing.Color.FromArgb(fr, fg, fb));
    }

    private void SetStatus(string msg) =>
        txtStatus.Text = $"{DateTime.Now:HH:mm:ss}  |  {msg}";
}

// ================================================================
//  Модель
// ================================================================
public class Employee
{
    public int    Id       { get; set; }
    public string Login    { get; set; } = "";
    public string Password { get; set; } = "";
    public string Role     { get; set; } = "";
    public string FullName { get; set; } = "";
    public string Email    { get; set; } = "";
}
