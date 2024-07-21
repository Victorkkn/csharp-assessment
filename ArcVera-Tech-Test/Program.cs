namespace ArcVera_Tech_Test
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        ///  Day 3 2024-07-20: finished the main quests, TODO: bonus and if had time optmize the excel export and paint just 5 first columns with red
        ///  Day 2 2024-07-19: created table from parquet and exported data to csv, problems with number rows from excel file
        ///  Day 1 2024-07-18: study parquet and recommended libraries, started to read data from parquet
        /// </summary>
        /// 

        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
            Application.Run(new frmMain());

            Task task = Execute.ExecuteCodeAsync();

            // Wait for the task to complete
            task.Wait();

        }       
    }
}