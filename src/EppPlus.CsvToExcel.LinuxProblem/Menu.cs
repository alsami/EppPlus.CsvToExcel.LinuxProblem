namespace EppPlus.CsvToExcel.LinuxProblem
{
    public class Menu
    {
        public string Main { get; }

        public string Desert { get; }

        public Menu(string main, string desert)
        {
            this.Main = main;
            this.Desert = desert;
        }
    }
}